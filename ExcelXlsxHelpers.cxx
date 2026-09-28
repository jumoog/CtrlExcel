#include <ExcelXlsxHelpers.hxx>

#include <BitVar.hxx>
#include <AnyTypeVar.hxx>
#include <DynVar.hxx>
#include <FloatVar.hxx>
#include <IntegerVar.hxx>
#include <LongVar.hxx>
#include <MappingVar.hxx>
#include <MixedVar.hxx>
#include <TextVar.hxx>
#include <TimeVar.hxx>

#include <cctype>
#include <climits>
#include <unordered_map>
#include <unordered_set>
#include <cmath>
#include <cstdio>
#include <cstring>
#include <ctime>
#include <memory>
#include <string>
#include <vector>

using namespace OpenXLSX;

namespace ExcelXlsxHelpers
{
  const Variable *unwrapAnyOrMixed(const Variable *val)
  {
    const Variable *current = val;

    while (current
        && (current->isA() == ANYTYPE_VAR || current->isA() == MIXED_VAR))
    {
      const AnyTypeVar *wrapped = static_cast<const AnyTypeVar *>(current);
      current = wrapped->getVar();
    }

    return current;
  }
}

namespace
{
  constexpr uint32_t EXCEL_FMT_DATE_TIME = 22;
  constexpr long long DAY_MILLIS = 24LL * 60LL * 60LL * 1000LL;
  constexpr long long DAY_SECONDS = 24LL * 60LL * 60LL;
  constexpr size_t EXCEL_MAX_CELL_CHARS = 32767;
  constexpr size_t EXCEL_MAX_SHEET_NAME_CHARS = 31;

  // Excel serial day of 1970-01-01. Serials below 61 precede Excel's
  // fictitious 1900-02-29 and are shifted by one day.
  constexpr long long EXCEL_UNIX_EPOCH_DAY = 25569;

  bool toLocalCalendarTime(time_t sec, std::tm &outTm)
  {
#ifdef _WIN32
    return localtime_s(&outTm, &sec) == 0;
#else
    return localtime_r(&sec, &outTm) != nullptr;
#endif
  }

  // Days since 1970-01-01 of a proleptic Gregorian date (H. Hinnant's
  // days_from_civil). Pure arithmetic, so it works outside the C runtime's
  // time_t range.
  long long daysFromCivil(long long y, unsigned m, unsigned d)
  {
    y -= m <= 2;
    const long long era = (y >= 0 ? y : y - 399) / 400;
    const unsigned yoe = static_cast<unsigned>(y - era * 400);
    const unsigned doy = (153 * (m > 2 ? m - 3 : m + 9) + 2) / 5 + d - 1;
    const unsigned doe = yoe * 365 + yoe / 4 - yoe / 100 + doy;
    return era * 146097 + static_cast<long long>(doe) - 719468;
  }

  // Inverse of daysFromCivil.
  void civilFromDays(long long z, long long &y, unsigned &m, unsigned &d)
  {
    z += 719468;
    const long long era = (z >= 0 ? z : z - 146096) / 146097;
    const unsigned doe = static_cast<unsigned>(z - era * 146097);
    const unsigned yoe = (doe - doe / 1460 + doe / 36524 - doe / 146096) / 365;
    const unsigned doy = doe - (365 * yoe + yoe / 4 - yoe / 100);
    const unsigned mp = (5 * doy + 2) / 153;
    d = doy - (153 * mp + 2) / 5 + 1;
    m = mp < 10 ? mp + 3 : mp - 9;
    y = static_cast<long long>(yoe) + era * 400 + (m <= 2);
  }

  long long excelDayToUnixDay(long long excelDay)
  {
    return excelDay - EXCEL_UNIX_EPOCH_DAY + (excelDay < 61 ? 1 : 0);
  }

  long long unixDayToExcelDay(long long unixDay)
  {
    long long excelDay = unixDay + EXCEL_UNIX_EPOCH_DAY;
    return excelDay < 61 ? excelDay - 1 : excelDay;
  }

  long long floorDiv(long long a, long long b)
  {
    long long q = a / b;
    return (a % b != 0 && (a < 0) != (b < 0)) ? q - 1 : q;
  }

  // UTC offset (seconds) for times the C runtime cannot convert: MSVC's
  // mktime/localtime reject times before 1970 and after 3000. Uses the zone's
  // offset in early January 1970 (standard time, no DST) in both directions,
  // so such times still round-trip.
  long long fallbackUtcOffset()
  {
    static const long long offset = []
    {
      const time_t probe = 2 * DAY_SECONDS; // a valid local time in every zone
      std::tm t{};
      if (!toLocalCalendarTime(probe, t))
        return 0LL;
      long long local = daysFromCivil(t.tm_year + 1900, static_cast<unsigned>(t.tm_mon + 1),
                                      static_cast<unsigned>(t.tm_mday)) * DAY_SECONDS
                      + t.tm_hour * 3600LL + t.tm_min * 60LL + t.tm_sec;
      return local - static_cast<long long>(probe);
    }();
    return offset;
  }

  // UTF-16 length of a UTF-8 string, which is what Excel's limits count.
  size_t utf16Length(const std::string &s)
  {
    size_t n = 0;
    for (unsigned char c : s)
    {
      if ((c & 0xC0) == 0x80)
        continue;               // continuation byte
      n += (c >= 0xF0) ? 2 : 1; // 4-byte sequences become surrogate pairs
    }
    return n;
  }

  // Cell text without the control characters XML 1.0 cannot represent (Excel
  // rejects the file otherwise). width receives the length in characters.
  // False if the text exceeds Excel's per-cell limit.
  bool toCellText(const char *text, std::string &out, size_t &width)
  {
    out.clear();
    for (const char *p = text; *p; ++p)
    {
      unsigned char c = static_cast<unsigned char>(*p);
      if (c < 0x20 && c != '\t' && c != '\n' && c != '\r')
        continue;
      out.push_back(*p);
    }
    width = utf16Length(out);
    return width <= EXCEL_MAX_CELL_CHARS;
  }

  size_t formattedWidth(const char *fmt, long long value)
  {
    char buf[32];
    int n = snprintf(buf, sizeof(buf), fmt, value);
    return n > 0 ? static_cast<size_t>(n) : 0;
  }

  XLStyleIndex createBoldHeaderFormat(XLDocument &doc)
  {
    XLStyleIndex boldFontIdx = doc.styles().fonts().create(
      doc.styles().fonts().fontByIndex(0));
    doc.styles().fonts().fontByIndex(boldFontIdx).setBold(true);

    XLStyleIndex boldFmtIdx = doc.styles().cellFormats().create(
      doc.styles().cellFormats().cellFormatByIndex(0));
    doc.styles().cellFormats().cellFormatByIndex(boldFmtIdx)
      .setFontIndex(boldFontIdx);
    doc.styles().cellFormats().cellFormatByIndex(boldFmtIdx)
      .setApplyFont(true);

    return boldFmtIdx;
  }

  XLStyleIndex createDateTimeFormat(XLDocument &doc)
  {
    XLStyleIndex dateTimeFmtIdx = doc.styles().cellFormats().create(
      doc.styles().cellFormats().cellFormatByIndex(0));
    doc.styles().cellFormats().cellFormatByIndex(dateTimeFmtIdx)
      .setNumberFormatId(EXCEL_FMT_DATE_TIME);
    doc.styles().cellFormats().cellFormatByIndex(dateTimeFmtIdx)
      .setApplyNumberFormat(true);

    return dateTimeFmtIdx;
  }

  //----------------------------------------------------------------------------
  // Date-format detection helpers
  //   OpenXLSX reports dates as XLValueType::Float. We inspect the cell's
  //   number-format to tell dates from plain numbers.
  //----------------------------------------------------------------------------

  // Built-in Excel number-format IDs that represent dates/times (ECMA-376).
  bool isBuiltinDateFormatId(unsigned int id)
  {
    return (id >= 14 && id <= 22)
        || (id >= 27 && id <= 36)
        || (id >= 45 && id <= 47);
  }

  // Scan a custom format-code string for date/time tokens (y m d h s)
  // while ignoring quoted literals, escaped chars and bracketed sections.
  bool isDateFormatCode(const std::string &code)
  {
    bool inQuote = false;
    bool inBracket = false;

    for (size_t i = 0; i < code.size(); i++)
    {
      char c = code[i];

      if (c == '"') { inQuote = !inQuote; continue; }
      if (inQuote) continue;
      if (c == '\\') { i++; continue; } // skip escaped char
      if (c == '[') { inBracket = true; continue; }
      if (c == ']') { inBracket = false; continue; }
      if (inBracket) continue;

      char lower = static_cast<char>(tolower(static_cast<unsigned char>(c)));
      if (lower == 'y' || lower == 'm' || lower == 'd'
       || lower == 'h' || lower == 's')
        return true;
    }
    return false;
  }

  //----------------------------------------------------------------------------
  // Read helpers
  //----------------------------------------------------------------------------

  // Whether the cell's number format is a date/time format. Results are
  // cached per cell-format index so each distinct format is inspected once.
  bool isDateCell(XLCell &cell, const XLStyles &styles,
                  std::unordered_map<XLStyleIndex, bool> &dateCache)
  {
    try
    {
      XLStyleIndex styleIdx = cell.cellFormat();
      auto it = dateCache.find(styleIdx);
      if (it != dateCache.end())
        return it->second;

      auto fmt = styles.cellFormats().cellFormatByIndex(styleIdx);
      unsigned int fmtId = fmt.numberFormatId();
      bool isDate = isBuiltinDateFormatId(fmtId);
      if (!isDate && fmtId >= 164)
      {
        std::string code = styles.numberFormats()
                                 .numberFormatById(fmtId)
                                 .formatCode();
        isDate = isDateFormatCode(code);
      }
      dateCache[styleIdx] = isDate;
      return isDate;
    }
    catch (...)
    {
      return false;
    }
  }

  // Convert an Excel date serial (local wall-clock time) to epoch seconds and
  // milliseconds; tm receives the calendar fields.
  time_t serialToLocalTime(double serial, std::tm &tm, PVSSshort &milli)
  {
    // Round the whole serial to milliseconds once and derive date, time of
    // day and milliseconds from it, so they cannot disagree by a second.
    long long totalMillis = llround(serial * static_cast<double>(DAY_MILLIS));
    long long excelDay    = totalMillis / DAY_MILLIS; // serial >= 1, never negative
    long long msOfDay     = totalMillis % DAY_MILLIS;
    long long unixDay     = excelDayToUnixDay(excelDay);

    long long year;
    unsigned month, day;
    civilFromDays(unixDay, year, month, day);

    tm = std::tm{};
    tm.tm_year  = static_cast<int>(year - 1900);
    tm.tm_mon   = static_cast<int>(month) - 1;
    tm.tm_mday  = static_cast<int>(day);
    tm.tm_hour  = static_cast<int>(msOfDay / 3600000LL);
    tm.tm_min   = static_cast<int>(msOfDay / 60000LL % 60LL);
    tm.tm_sec   = static_cast<int>(msOfDay / 1000LL % 60LL);
    // Serials carry no UTC offset: in the repeated hour after the DST
    // fall-back, mktime has to pick one of the two instants.
    tm.tm_isdst = -1;

    milli = static_cast<PVSSshort>(msOfDay % 1000LL);

    // Mirrors localTimeToSerial: the C runtime from 1970 on (with DST), the
    // fixed fallback offset where it fails.
    std::tm probe = tm;
    time_t sec = (unixDay >= 2) ? mktime(&probe) : static_cast<time_t>(-1);
    if (sec == static_cast<time_t>(-1))
      sec = static_cast<time_t>(unixDay * DAY_SECONDS + msOfDay / 1000LL - fallbackUtcOffset());
    return sec;
  }

  // XLDateTime rejects serials below 1.0 (time-only values such as 08:00),
  // so those are treated as plain numbers.
  bool isDateValue(const XLCellValue &val, XLCell &cell, const XLStyles &styles,
                   std::unordered_map<XLStyleIndex, bool> &dateCache)
  {
    return val.type() == XLValueType::Float
        && val.get<double>() >= 1.0
        && isDateCell(cell, styles, dateCache);
  }

  // Header cells may hold numbers, dates or booleans; get<std::string> throws
  // for those, and getString() renders 2024 as "2024.000000".
  std::string headerText(XLCell &cell, const XLStyles &styles,
                         std::unordered_map<XLStyleIndex, bool> &dateCache)
  {
    XLCellValue val = cell.value();

    if (isDateValue(val, cell, styles, dateCache))
    {
      std::tm tm{};
      PVSSshort milli = 0;
      serialToLocalTime(val.get<double>(), tm, milli);

      bool hasTime = tm.tm_hour != 0 || tm.tm_min != 0 || tm.tm_sec != 0;
      char buf[32];
      strftime(buf, sizeof(buf), hasTime ? "%Y-%m-%d %H:%M:%S" : "%Y-%m-%d", &tm);
      return buf;
    }

    switch (val.type())
    {
      case XLValueType::String:
        return val.get<std::string>();
      case XLValueType::Integer:
        return std::to_string(val.get<int64_t>());
      case XLValueType::Float:
      {
        double d = val.get<double>();
        double intpart;
        if (modf(d, &intpart) == 0.0 && fabs(intpart) < 1e15)
          return std::to_string(static_cast<long long>(intpart));
        char buf[32];
        snprintf(buf, sizeof(buf), "%.15g", d);
        return buf;
      }
      case XLValueType::Boolean:
        return val.get<bool>() ? "TRUE" : "FALSE";
      default:
        return std::string();
    }
  }

  // Set a mapping value using the cell type reported by OpenXLSX.
  void setTypedCellCached(MappingVar &row, const Variable &key,
                          XLCell &cell,
                          const XLStyles &styles,
                          std::unordered_map<XLStyleIndex, bool> &dateCache)
  {
    XLCellValue val = cell.value();
    auto type = val.type();

    switch (type)
    {
      case XLValueType::Empty:
        row.setAt(key, TextVar(""));
        return;

      case XLValueType::Boolean:
        row.setAt(key, BitVar(val.get<bool>()));
        return;

      case XLValueType::Integer:
      {
        int64_t ival = val.get<int64_t>();
        if (ival >= INT_MIN && ival <= INT_MAX)
          row.setAt(key, IntegerVar(static_cast<int>(ival)));
        else
          row.setAt(key, LongVar(ival));
        return;
      }

      case XLValueType::Float:
      {
        double dval = val.get<double>();
        if (isDateValue(val, cell, styles, dateCache))
        {
          std::tm tm{};
          PVSSshort milli = 0;
          time_t sec = serialToLocalTime(dval, tm, milli);
          row.setAt(key, TimeVar(sec, milli));
        }
        else
        {
          double intpart;
          if (modf(dval, &intpart) == 0.0
           && intpart >= INT_MIN && intpart <= INT_MAX)
          {
            row.setAt(key, IntegerVar(static_cast<int>(intpart)));
          }
          else
          {
            row.setAt(key, FloatVar(dval));
          }
        }
        return;
      }

      case XLValueType::String:
        row.setAt(key, TextVar(val.get<std::string>().c_str()));
        return;

      default:
        row.setAt(key, TextVar(""));
        return;
    }
  }

  //----------------------------------------------------------------------------
  // Write helpers
  //----------------------------------------------------------------------------

  // The row as a mapping (unwrapping anytype/mixed), or nullptr if it is not one.
  const MappingVar *asMapping(const Variable *rowVar)
  {
    const Variable *inner = ExcelXlsxHelpers::unwrapAnyOrMixed(rowVar);
    return (inner && inner->isA() == MAPPING_VAR)
      ? static_cast<const MappingVar *>(inner)
      : nullptr;
  }

  // Epoch seconds and milliseconds to the Excel serial of the local
  // wall-clock time (Excel serials have no timezone). Inverse of
  // serialToLocalTime.
  double localTimeToSerial(time_t sec, PVSSshort milli)
  {
    long long unixDay;
    long long secOfDay;
    std::tm lt{};
    if (sec >= 0 && toLocalCalendarTime(sec, lt))
    {
      unixDay  = daysFromCivil(lt.tm_year + 1900, static_cast<unsigned>(lt.tm_mon + 1),
                               static_cast<unsigned>(lt.tm_mday));
      secOfDay = lt.tm_hour * 3600LL + lt.tm_min * 60LL + lt.tm_sec;
    }
    else
    {
      long long wall = static_cast<long long>(sec) + fallbackUtcOffset();
      unixDay  = floorDiv(wall, DAY_SECONDS);
      secOfDay = wall - unixDay * DAY_SECONDS;
    }

    return static_cast<double>(unixDayToExcelDay(unixDay))
         + (static_cast<double>(secOfDay) * 1000.0 + milli) / static_cast<double>(DAY_MILLIS);
  }

  // Write a single WinCC OA Variable to an OpenXLSX cell. width receives the
  // displayed length for column sizing. False if text exceeds Excel's limit.
  bool writeTypedCell(XLCell &cell, const Variable *val,
                      XLStyleIndex dateTimeFmtIdx, size_t &width)
  {
    width = 0;
    if (!val)
    {
      cell.value() = std::string();
      return true;
    }

    switch (val->isA())
    {
      case INTEGER_VAR:
      {
        long long v = static_cast<const IntegerVar *>(val)->getValue();
        cell.value() = static_cast<int64_t>(v);
        width = formattedWidth("%lld", v);
        return true;
      }
      case LONG_VAR:
      {
        long long v = static_cast<const LongVar *>(val)->getValue();
        cell.value() = static_cast<int64_t>(v);
        width = formattedWidth("%lld", v);
        return true;
      }
      case FLOAT_VAR:
      {
        double v = static_cast<const FloatVar *>(val)->getValue();
        cell.value() = v;
        char buf[32];
        int n = snprintf(buf, sizeof(buf), "%.15g", v);
        width = n > 0 ? static_cast<size_t>(n) : 0;
        return true;
      }
      case BIT_VAR:
        cell.value() = static_cast<const BitVar *>(val)->isTrue();
        width = 5; // "FALSE"
        return true;
      case TIME_VAR:
      {
        const TimeVar *timeVal = static_cast<const TimeVar *>(val);
        cell.value() = XLDateTime(localTimeToSerial(
          static_cast<time_t>(timeVal->getSeconds()), timeVal->getMilli()));
        cell.setCellFormat(dateTimeFmtIdx);
        width = 19; // "YYYY-MM-DD hh:mm:ss"
        return true;
      }
      case TEXT_VAR:
      {
        std::string text;
        if (!toCellText(static_cast<const TextVar *>(val)->getValue(), text, width))
          return false;
        cell.value() = text;
        return true;
      }
      case ANYTYPE_VAR:
      case MIXED_VAR:
      {
        const Variable *inner = ExcelXlsxHelpers::unwrapAnyOrMixed(val);
        if (!inner)
        {
          cell.value() = std::string();
          return true;
        }

        return writeTypedCell(cell, inner, dateTimeFmtIdx, width);
      }
      default:
      {
        CharString str = val->formatValue(CharString());
        std::string text;
        if (!toCellText(str.c_str(), text, width))
          return false;
        cell.value() = text;
        return true;
      }
    }
  }
} // namespace

namespace ExcelXlsxHelpers
{
  // Read an open worksheet into a DynVar of MappingVar rows.
  void readSheetRows(XLWorksheet &wks, XLDocument &doc,
                     DynVar &result, bool useHeaders, bool skipHidden)
  {
    uint32_t rowCount = wks.rowCount();
    uint16_t colCount = wks.columnCount();

    if (rowCount == 0 || colCount == 0)
      return;

    // Pre-fetch styles once and cache format-index → is_date results so each
    // unique cell format is inspected only once across the entire sheet.
    XLStyles styles = doc.styles();
    std::unordered_map<XLStyleIndex, bool> dateCache;

    // One mapping key per column: the header text, or the 1-based column
    // number when headers are off or a header cell is empty. Repeated
    // headers get a numeric suffix ("Value", "Value_2") so no two columns
    // share a key and overwrite each other.
    std::vector<std::unique_ptr<Variable>> keys;
    keys.reserve(colCount);
    uint32_t dataStartRow = 1;

    if (useHeaders)
    {
      std::unordered_set<std::string> seen;
      uint16_t c = 1;
      for (auto& cell : wks.row(1).cells(colCount))
      {
        std::string name = headerText(cell, styles, dateCache);
        if (name.empty())
        {
          keys.emplace_back(new IntegerVar(c));
        }
        else
        {
          std::string unique = name;
          for (int n = 2; !seen.insert(unique).second; ++n)
            unique = name + "_" + std::to_string(n);
          keys.emplace_back(new TextVar(unique.c_str()));
        }
        ++c;
      }
      dataStartRow = 2;
    }

    for (uint16_t c = static_cast<uint16_t>(keys.size()) + 1; c <= colCount; ++c)
      keys.emplace_back(new IntegerVar(c));

    // Header-only sheet: rows(2, 1) would still yield one bogus row.
    if (dataStartRow > rowCount)
      return;

    for (auto& xlRow : wks.rows(dataStartRow, rowCount))
    {
      if (skipHidden && xlRow.isHidden())
        continue;

      // cells(colCount) yields exactly colCount cells, one per key.
      MappingVar rowMap;
      size_t c = 0;
      for (auto& cell : xlRow.cells(colCount))
        setTypedCellCached(rowMap, *keys[c++], cell, styles, dateCache);
      result.append(rowMap);
    }
  }

  // Write a DynVar of MappingVars to an OpenXLSX worksheet. Columns are the
  // keys of all rows in order of first appearance; row 1 holds them as bold
  // headers.
  bool writeSheetData(XLWorksheet &wks, const DynVar &data, XLDocument &doc,
                      std::string &error)
  {
    unsigned int numRows = data.getNumberOfItems();
    if (numRows == 0)
      return true;

    // Validate every row before touching the sheet.
    std::vector<const MappingVar *> rows;
    rows.reserve(numRows);
    for (unsigned int r = 0; r < numRows; r++)
    {
      const MappingVar *row = asMapping(data.getAt(r));
      if (!row)
      {
        error = "row " + std::to_string(r + 1) + " is not a mapping";
        return false;
      }
      rows.push_back(row);
    }

    // Union of all keys, so keys missing from the first row are not dropped.
    // Lookups use the original key Variables: stringified TextVar keys would
    // miss non-text keys such as the int column keys readSheetRows produces.
    MappingVar seenKeys;
    std::vector<const Variable *> keys;
    for (const MappingVar *row : rows)
    {
      for (unsigned int k = 0; k < row->getNumberOfItems(); k++)
      {
        const Variable *key = row->getKey(k);
        if (!seenKeys.getAt(*key))
        {
          seenKeys.setAt(*key, BitVar(true));
          keys.push_back(key);
        }
      }
    }

    size_t numCols = keys.size();
    if (numCols == 0)
      return true;

    // Create bold cell format for the header row.
    // XLFont/XLCellFormat are XML-node proxies: call create() first to duplicate
    // the default entry, then modify the new entry so the default is not mutated.
    XLStyleIndex boldFmtIdx = createBoldHeaderFormat(doc);

    // Built-in format 22: date + time display. Without a date format, Excel
    // shows the serial value (e.g. 46105.837...).
    XLStyleIndex dateTimeFmtIdx = createDateTimeFormat(doc);

    // Max displayed width per column, starting with the header.
    std::vector<size_t> maxWidths(numCols);

    // Write column headers in row 1 (bold)
    for (size_t c = 0; c < numCols; c++)
    {
      CharString name = keys[c]->formatValue(CharString());
      std::string text;
      if (!toCellText(name.c_str(), text, maxWidths[c]))
      {
        error = "column header exceeds " + std::to_string(EXCEL_MAX_CELL_CHARS) + " characters";
        return false;
      }

      auto cell = wks.cell(1, static_cast<uint16_t>(c + 1));
      cell.value() = text;
      cell.setCellFormat(boldFmtIdx);
    }

    // Write data rows starting from row 2
    for (unsigned int r = 0; r < numRows; r++)
    {
      for (size_t c = 0; c < numCols; c++)
      {
        const Variable *cellVal = rows[r]->getAt(*keys[c]);
        auto cell = wks.cell(
          static_cast<uint32_t>(r + 2),
          static_cast<uint16_t>(c + 1)
        );

        size_t width = 0;
        if (!writeTypedCell(cell, cellVal, dateTimeFmtIdx, width))
        {
          error = "row " + std::to_string(r + 1) + ", column "
                + std::to_string(c + 1) + ": text exceeds "
                + std::to_string(EXCEL_MAX_CELL_CHARS) + " characters";
          return false;
        }
        if (width > maxWidths[c])
          maxWidths[c] = width;
      }
    }

    // Set column widths (character width + padding)
    for (size_t c = 0; c < numCols; c++)
      wks.column(static_cast<uint16_t>(c + 1))
          .setWidth(static_cast<float>(maxWidths[c]) + 2.0f);

    return true;
  }

  std::string checkSheetNames(const std::vector<std::string> &names)
  {
    std::unordered_set<std::string> seen;
    for (const std::string &name : names)
    {
      std::string quoted = "sheet name '" + name + "'";
      if (name.empty())
        return "sheet name must not be empty";
      if (utf16Length(name) > EXCEL_MAX_SHEET_NAME_CHARS)
        return quoted + " is longer than " + std::to_string(EXCEL_MAX_SHEET_NAME_CHARS) + " characters";
      for (unsigned char c : name)
      {
        if (c < 0x20 || strchr("\\/?*[]:", c))
          return quoted + " contains a character Excel does not allow (\\ / ? * [ ] : or a control character)";
      }
      if (name.front() == '\'' || name.back() == '\'')
        return quoted + " must not start or end with an apostrophe";

      // Excel compares sheet names case-insensitively (ASCII folding here).
      std::string lower = name;
      for (char &ch : lower)
        ch = static_cast<char>(tolower(static_cast<unsigned char>(ch)));
      if (lower == "history")
        return quoted + " is reserved by Excel";
      if (!seen.insert(lower).second)
        return quoted + " duplicates another sheet name (Excel ignores case)";
    }
    return std::string();
  }
}
