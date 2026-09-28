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

  bool toLocalCalendarTime(time_t sec, std::tm &outTm)
  {
#ifdef _WIN32
    return localtime_s(&outTm, &sec) == 0;
#else
    return localtime_r(&sec, &outTm) != nullptr;
#endif
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
  // milliseconds; tm receives the normalised calendar fields.
  time_t serialToLocalTime(double serial, std::tm &tm, PVSSshort &milli)
  {
    // Round the whole serial to milliseconds once and derive both the time
    // of day and the millisecond part from it. XLDateTime::tm() truncates, so
    // taking seconds from it and milliseconds from a separately rounded
    // fraction can be off by one second.
    long long totalMillis = llround(serial * static_cast<double>(DAY_MILLIS));
    long long roundedDay  = totalMillis / DAY_MILLIS;
    long long msOfDay     = totalMillis % DAY_MILLIS;

    // tm() yields the calendar date of floor(serial); carry into the next day
    // when rounding crossed midnight (mktime normalises).
    tm = XLDateTime(serial).tm();
    tm.tm_mday += static_cast<int>(roundedDay - static_cast<long long>(floor(serial)));
    tm.tm_hour  = static_cast<int>(msOfDay / 3600000LL);
    tm.tm_min   = static_cast<int>(msOfDay / 60000LL % 60LL);
    tm.tm_sec   = static_cast<int>(msOfDay / 1000LL % 60LL);
    // Serials carry no UTC offset: in the repeated hour after the DST
    // fall-back, mktime has to pick one of the two instants.
    tm.tm_isdst = -1;

    milli = static_cast<PVSSshort>(msOfDay % 1000LL);
    return mktime(&tm);
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

  // Write a single WinCC OA Variable to an OpenXLSX cell.
  void writeTypedCell(XLCell &cell, const Variable *val,
                      XLStyleIndex dateTimeFmtIdx)
  {
    if (!val)
    {
      cell.value() = std::string();
      return;
    }

    switch (val->isA())
    {
      case INTEGER_VAR:
        cell.value() = static_cast<int64_t>(
          static_cast<const IntegerVar *>(val)->getValue());
        return;
      case LONG_VAR:
        cell.value() = static_cast<int64_t>(
          static_cast<const LongVar *>(val)->getValue());
        return;
      case FLOAT_VAR:
        cell.value() = static_cast<const FloatVar *>(val)->getValue();
        return;
      case BIT_VAR:
        cell.value() = static_cast<const BitVar *>(val)->isTrue();
        return;
      case TIME_VAR:
      {
        const TimeVar *timeVal = static_cast<const TimeVar *>(val);
        time_t sec = static_cast<time_t>(timeVal->getSeconds());

        // Excel date/time serials have no timezone. Convert epoch seconds to
        // local calendar fields first, then write those fields as XLDateTime
        // so displayed wall-clock time matches WinCC OA local time.
        std::tm localTm{};
        XLDateTime dt = toLocalCalendarTime(sec, localTm)
          ? XLDateTime(localTm)
          : XLDateTime(sec);

        // XLDateTime only carries whole seconds; add the milliseconds to the
        // serial so they survive a round-trip.
        cell.value() = XLDateTime(dt.serial()
          + static_cast<double>(timeVal->getMilli()) / static_cast<double>(DAY_MILLIS));

        cell.setCellFormat(dateTimeFmtIdx);
        return;
      }
      case TEXT_VAR:
        cell.value() = std::string(
          static_cast<const TextVar *>(val)->getValue());
        return;
      case ANYTYPE_VAR:
      case MIXED_VAR:
      {
        const Variable *inner = ExcelXlsxHelpers::unwrapAnyOrMixed(val);
        if (!inner)
        {
          cell.value() = std::string();
          return;
        }

        writeTypedCell(cell, inner, dateTimeFmtIdx);
        return;
      }
      default:
      {
        CharString str = val->formatValue(CharString());
        cell.value() = std::string(str.c_str());
        return;
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

  // Write a DynVar of MappingVars to an OpenXLSX worksheet.
  // Column headers are taken from the keys of the first mapping row.
  bool writeSheetData(XLWorksheet &wks, const DynVar &data, XLDocument &doc)
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
        return false;
      rows.push_back(row);
    }

    const MappingVar &firstRow = *rows[0];
    unsigned int numCols = firstRow.getNumberOfItems();
    if (numCols == 0)
      return true;

    // Keep the first row's keys for lookups: stringifying them and looking up
    // by TextVar would miss non-text keys such as the integer column keys
    // readSheetRows produces without headers.
    std::vector<const Variable *> keys;
    std::vector<CharString> columnNames;
    keys.reserve(numCols);
    columnNames.reserve(numCols);
    for (unsigned int c = 0; c < numCols; c++)
    {
      const Variable *key = firstRow.getKey(c);
      keys.push_back(key);
      columnNames.push_back(key->formatValue(CharString()));
    }

    // Track max character width per column (init with header lengths)
    std::vector<size_t> maxWidths(numCols);
    for (unsigned int c = 0; c < numCols; c++)
      maxWidths[c] = strlen(columnNames[c].c_str());

    // Create bold cell format for the header row.
    // XLFont/XLCellFormat are XML-node proxies: call create() first to duplicate
    // the default entry, then modify the new entry so the default is not mutated.
    XLStyleIndex boldFmtIdx = createBoldHeaderFormat(doc);

    // Built-in format 22: date + time display. Without a date format, Excel
    // shows the serial value (e.g. 46105.837...).
    XLStyleIndex dateTimeFmtIdx = createDateTimeFormat(doc);

    // Write column headers in row 1 (bold)
    for (unsigned int c = 0; c < numCols; c++)
    {
      auto cell = wks.cell(1, static_cast<uint16_t>(c + 1));
      cell.value() = std::string(columnNames[c].c_str());
      cell.setCellFormat(boldFmtIdx);
    }

    // Write data rows starting from row 2
    for (unsigned int r = 0; r < numRows; r++)
    {
      for (unsigned int c = 0; c < numCols; c++)
      {
        const Variable *cellVal = rows[r]->getAt(*keys[c]);
        auto cell = wks.cell(
          static_cast<uint32_t>(r + 2),
          static_cast<uint16_t>(c + 1)
        );
        writeTypedCell(cell, cellVal, dateTimeFmtIdx);

        if (cellVal)
        {
          CharString str = cellVal->formatValue(CharString());
          size_t len = strlen(str.c_str());
          if (len > maxWidths[c])
            maxWidths[c] = len;
        }
      }
    }

    // Set column widths (character width + padding)
    for (unsigned int c = 0; c < numCols; c++)
      wks.column(static_cast<uint16_t>(c + 1))
          .setWidth(static_cast<float>(maxWidths[c]) + 2.0f);

    return true;
  }
}
