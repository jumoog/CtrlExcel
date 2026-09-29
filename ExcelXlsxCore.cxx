#include <ExcelXlsxCore.hxx>

#include <cctype>
#include <cerrno>
#include <cmath>
#include <cstdio>
#include <cstdlib>
#include <cstring>
#include <limits>
#include <sstream>
#include <unordered_set>

using namespace OpenXLSX;

namespace
{
  constexpr size_t EXCEL_MAX_SHEET_NAME_CHARS = 31;

  bool isHex(unsigned char c)
  {
    return isxdigit(c) != 0;
  }

  // Whether p starts an Excel escape sequence "_xHHHH_".
  bool isExcelEscape(const unsigned char *p)
  {
    return p[0] == '_' && (p[1] == 'x' || p[1] == 'X')
        && isHex(p[2]) && isHex(p[3]) && isHex(p[4]) && isHex(p[5]) && p[6] == '_';
  }

  void appendExcelEscape(std::string &out, unsigned int codeUnit)
  {
    char buf[8];
    snprintf(buf, sizeof(buf), "_x%04X_", codeUnit);
    out += buf;
  }

  bool toLocalCalendarTime(time_t sec, std::tm &outTm)
  {
#ifdef _WIN32
    return localtime_s(&outTm, &sec) == 0;
#else
    return localtime_r(&sec, &outTm) != nullptr;
#endif
  }

  // Seconds east of UTC at the given local calendar time, or OFFSET_VARIES if
  // mktime cannot convert it.
  long long utcOffsetAt(std::tm tm, long long unixDay, long long secOfDay)
  {
    tm.tm_hour  = static_cast<int>(secOfDay / 3600);
    tm.tm_min   = static_cast<int>(secOfDay / 60 % 60);
    tm.tm_sec   = static_cast<int>(secOfDay % 60);
    tm.tm_isdst = -1;
    time_t t = mktime(&tm);
    if (t == static_cast<time_t>(-1))
      return ExcelXlsxCore::ReadState::OFFSET_VARIES;
    return unixDay * ExcelXlsxCore::DAY_SECONDS + secOfDay - static_cast<long long>(t);
  }
}

namespace ExcelXlsxCore
{
  long long daysFromCivil(long long y, unsigned m, unsigned d)
  {
    y -= m <= 2;
    const long long era = (y >= 0 ? y : y - 399) / 400;
    const unsigned yoe = static_cast<unsigned>(y - era * 400);
    const unsigned doy = (153 * (m > 2 ? m - 3 : m + 9) + 2) / 5 + d - 1;
    const unsigned doe = yoe * 365 + yoe / 4 - yoe / 100 + doy;
    return era * 146097 + static_cast<long long>(doe) - 719468;
  }

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

  std::string toXmlUtf8(const char *text, bool cellText)
  {
    std::string out;
    const unsigned char *p = reinterpret_cast<const unsigned char *>(text);
    while (*p)
    {
      const unsigned char c = *p;
      size_t len = (c < 0x80)               ? 1
                 : (c >= 0xC2 && c <= 0xDF) ? 2
                 : (c >= 0xE0 && c <= 0xEF) ? 3
                 : (c >= 0xF0 && c <= 0xF4) ? 4
                 : 0;
      bool valid = len > 0;
      for (size_t i = 1; valid && i < len; i++)
        valid = (p[i] & 0xC0) == 0x80; // also stops at the terminating NUL

      // Reject overlong forms, UTF-16 surrogates and code points > U+10FFFF.
      if (valid && len == 3
       && ((c == 0xE0 && p[1] < 0xA0) || (c == 0xED && p[1] >= 0xA0)))
        valid = false;
      if (valid && len == 4
       && ((c == 0xF0 && p[1] < 0x90) || (c == 0xF4 && p[1] >= 0x90)))
        valid = false;

      if (!valid)
      {
        out.push_back(static_cast<char>(0xC0 | (c >> 6)));
        out.push_back(static_cast<char>(0x80 | (c & 0x3F)));
        ++p;
        continue;
      }

      if (cellText)
      {
        if (len == 1 && c < 0x20 && c != '\t' && c != '\n')
        {
          appendExcelEscape(out, c);
          ++p;
          continue;
        }
        if (len == 3 && c == 0xEF && p[1] == 0xBF && p[2] >= 0xBE)
        {
          appendExcelEscape(out, 0xFFC0u | (p[2] & 0x3Fu)); // U+FFFE / U+FFFF
          p += 3;
          continue;
        }
        if (isExcelEscape(p))
        {
          out += "_x005F_"; // the '_' itself; "xHHHH_" follows as plain text
          ++p;
          continue;
        }
      }

      out.append(reinterpret_cast<const char *>(p), len);
      p += len;
    }
    return out;
  }

  std::string decodeExcelEscapes(const std::string &text)
  {
    if (text.find("_x") == std::string::npos && text.find("_X") == std::string::npos)
      return text;

    std::string out;
    const unsigned char *p = reinterpret_cast<const unsigned char *>(text.c_str());
    while (*p)
    {
      if (isExcelEscape(p))
      {
        unsigned long cu = strtoul(std::string(reinterpret_cast<const char *>(p + 2), 4).c_str(), nullptr, 16);
        if (cu < 0x20 || cu == 0x5F)
        {
          if (cu != 0)
            out.push_back(static_cast<char>(cu));
          p += 7;
          continue;
        }
        if (cu == 0xFFFE || cu == 0xFFFF)
        {
          out.push_back(static_cast<char>(0xEF));
          out.push_back(static_cast<char>(0xBF));
          out.push_back(static_cast<char>(0x80 | (cu & 0x3F)));
          p += 7;
          continue;
        }
      }
      out.push_back(static_cast<char>(*p++));
    }
    return out;
  }

  std::u32string foldCase(const std::string &utf8)
  {
    std::u32string out;
    const unsigned char *p = reinterpret_cast<const unsigned char *>(utf8.c_str());
    while (*p)
    {
      char32_t cp;
      if (*p < 0x80)      { cp = *p++; }
      else if (*p < 0xE0) { cp = (p[0] & 0x1Fu) << 6 | (p[1] & 0x3Fu); p += 2; }
      else if (*p < 0xF0) { cp = (p[0] & 0x0Fu) << 12 | (p[1] & 0x3Fu) << 6 | (p[2] & 0x3Fu); p += 3; }
      else                { cp = (p[0] & 0x07u) << 18 | (p[1] & 0x3Fu) << 12 | (p[2] & 0x3Fu) << 6 | (p[3] & 0x3Fu); p += 4; }

      if ((cp >= U'A' && cp <= U'Z') || (cp >= 0xC0 && cp <= 0xDE && cp != 0xD7)
       || (cp >= 0x391 && cp <= 0x3AB && cp != 0x3A2) || (cp >= 0x410 && cp <= 0x42F))
        cp += 0x20;
      else if (cp >= 0x400 && cp <= 0x40F)
        cp += 0x50;
      else if (((cp >= 0x100 && cp <= 0x137) || (cp >= 0x14A && cp <= 0x177)) && cp % 2 == 0)
        cp += 1;
      else if (((cp >= 0x139 && cp <= 0x148) || (cp >= 0x179 && cp <= 0x17E)) && cp % 2 == 1)
        cp += 1;

      out.push_back(cp);
    }
    return out;
  }

  size_t numberWidth(double value)
  {
    char buf[32];
    int n = snprintf(buf, sizeof(buf), "%.15g", value);
    return n > 0 ? static_cast<size_t>(n) : 0;
  }

  //----------------------------------------------------------------------------
  // Cell formats
  //----------------------------------------------------------------------------

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
  // Date-format detection
  //   OpenXLSX reports dates as XLValueType::Float. We inspect the cell's
  //   number-format to tell dates from plain numbers.
  //----------------------------------------------------------------------------

  bool isBuiltinDateFormatId(unsigned int id)
  {
    return (id >= 14 && id <= 22)
        || (id >= 27 && id <= 36)
        || id == 45 || id == 47;
  }

  bool isDateFormatCode(const std::string &code)
  {
    bool inQuote = false;
    bool inBracket = false;
    std::string bracket;

    for (size_t i = 0; i < code.size(); i++)
    {
      char c = code[i];

      if (c == '"') { inQuote = !inQuote; continue; }
      if (inQuote) continue;
      if (c == '\\') { i++; continue; } // skip escaped char
      if (c == '[') { inBracket = true; bracket.clear(); continue; }
      if (c == ']')
      {
        inBracket = false;
        if (!bracket.empty()
         && bracket.find_first_not_of("hHmMsS") == std::string::npos)
          return false;
        continue;
      }
      if (inBracket) { bracket.push_back(c); continue; }

      char lower = static_cast<char>(tolower(static_cast<unsigned char>(c)));
      if (lower == 'y' || lower == 'm' || lower == 'd'
       || lower == 'h' || lower == 's')
        return true;
    }
    return false;
  }

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

  //----------------------------------------------------------------------------
  // Dates
  //----------------------------------------------------------------------------

  LocalTime serialToLocalTime(double serial, std::tm &tm, int16_t &milli, time_t &sec,
                              time_t maxSec, ReadState &state)
  {
    // Round the whole serial to milliseconds once and derive date, time of
    // day and milliseconds from it, so they cannot disagree by a second.
    long long totalMillis = llround(serial * static_cast<double>(DAY_MILLIS));
    long long excelDay    = totalMillis / DAY_MILLIS; // serial >= 1, never negative
    long long msOfDay     = totalMillis % DAY_MILLIS;
    long long unixDay     = excelDayToUnixDay(excelDay);
    long long secOfDay    = msOfDay / 1000LL;

    long long year;
    unsigned month, day;
    civilFromDays(unixDay, year, month, day);

    tm = std::tm{};
    tm.tm_year  = static_cast<int>(year - 1900);
    tm.tm_mon   = static_cast<int>(month) - 1;
    tm.tm_mday  = static_cast<int>(day);
    tm.tm_hour  = static_cast<int>(secOfDay / 3600);
    tm.tm_min   = static_cast<int>(secOfDay / 60 % 60);
    tm.tm_sec   = static_cast<int>(secOfDay % 60);
    tm.tm_isdst = -1;

    milli = static_cast<int16_t>(msOfDay % 1000LL);

    sec = 0;
    if (unixDay < 0)
      return LocalTime::OutOfRange; // before 1970 even in local time

    auto cached = state.dayOffset.find(unixDay);
    if (cached == state.dayOffset.end())
    {
      long long atStart = utcOffsetAt(tm, unixDay, 0);
      long long atEnd   = utcOffsetAt(tm, unixDay, DAY_SECONDS - 1);
      cached = state.dayOffset.emplace(
        unixDay, atStart == atEnd ? atStart : ReadState::OFFSET_VARIES).first;
    }

    LocalTime result = LocalTime::Ok;
    long long t;
    if (cached->second != ReadState::OFFSET_VARIES)
    {
      t = unixDay * DAY_SECONDS + secOfDay - cached->second;
    }
    else
    {
      // DST transition day (or a day mktime cannot convert). Serials carry no
      // UTC offset: in the repeated hour after the fall-back mktime has to
      // pick one of the two instants, and a time in the hour skipped at the
      // spring-forward does not exist and is moved by mktime.
      std::tm probe = tm;
      time_t converted = mktime(&probe);
      if (converted == static_cast<time_t>(-1))
        return LocalTime::OutOfRange; // mktime fails before the epoch (and on MSVC after 3000)
      if (probe.tm_hour != tm.tm_hour || probe.tm_min != tm.tm_min)
        result = LocalTime::Shifted;
      t = static_cast<long long>(converted);
    }

    if (t < 0 || t > static_cast<long long>(maxSec))
      return LocalTime::OutOfRange;

    sec = static_cast<time_t>(t);
    return result;
  }

  bool localTimeToSerial(time_t sec, int milli, double &serial)
  {
    std::tm lt{};
    if (!toLocalCalendarTime(sec, lt))
      return false;

    long long unixDay = daysFromCivil(lt.tm_year + 1900, static_cast<unsigned>(lt.tm_mon + 1),
                                      static_cast<unsigned>(lt.tm_mday));
    long long secOfDay = lt.tm_hour * 3600LL + lt.tm_min * 60LL + lt.tm_sec;

    serial = static_cast<double>(unixDayToExcelDay(unixDay))
           + (static_cast<double>(secOfDay) * 1000.0 + milli) / static_cast<double>(DAY_MILLIS);
    return true;
  }

  //----------------------------------------------------------------------------
  // Cell values
  //----------------------------------------------------------------------------

  std::string rawNumberText(const XLCell &cell)
  {
    // Reused: constructing a stream per cell dominated the cost.
    static thread_local std::ostringstream xml;
    xml.str(std::string());
    xml.clear();
    cell.print(xml);
    const std::string node = xml.str();
    size_t begin = node.find("<v>");
    size_t end = node.find("</v>");
    if (begin == std::string::npos || end == std::string::npos || end < begin)
      return std::string();
    return node.substr(begin + 3, end - begin - 3);
  }

  bool cellNumber(const XLCell &cell, const XLCellValue &val, CellNumber &out)
  {
    if (val.type() == XLValueType::Float)
    {
      out.value = val.get<double>();
      return true;
    }
    if (val.type() != XLValueType::Integer)
      return false;

    out.integer = val.get<int64_t>();
    out.value = static_cast<double>(out.integer);
    out.isInteger = true;

    // OpenXLSX parses Integer text with as_llong, which stops at an exponent
    // ("1E+20" becomes 1) and clamps values beyond int64 to its limits.
    // Exponent notation always has a one-digit mantissa, so only results
    // -9..9 and the two limits need re-parsing from the XML.
    bool maybeExponent = out.integer >= -9 && out.integer <= 9;
    bool maybeOverflow = out.integer == INT64_MAX || out.integer == INT64_MIN;
    if (maybeExponent || maybeOverflow)
    {
      std::string text = rawNumberText(cell);
      bool reparse = text.find_first_of("eE") != std::string::npos;
      if (!reparse && maybeOverflow)
      {
        errno = 0;
        strtoll(text.c_str(), nullptr, 10);
        reparse = errno == ERANGE;
      }
      if (reparse)
      {
        out.value = strtod(text.c_str(), nullptr);
        out.isInteger = false;
      }
    }
    return true;
  }

  bool isDateNumber(const CellNumber &num, XLCell &cell, const XLStyles &styles,
                    std::unordered_map<XLStyleIndex, bool> &dateCache)
  {
    return num.value >= 1.0 && isDateCell(cell, styles, dateCache);
  }

  bool isUncachedFormula(const XLCell &cell, const XLCellValue &val)
  {
    return cell.hasFormula()
        && val.type() != XLValueType::String
        && rawNumberText(cell).empty();
  }

  std::string headerText(XLCell &cell, const XLStyles &styles,
                         std::unordered_map<XLStyleIndex, bool> &dateCache,
                         ReadState &state)
  {
    XLCellValue val = cell.value();

    if (isUncachedFormula(cell, val))
      return std::string();

    CellNumber num;
    if (cellNumber(cell, val, num))
    {
      if (isDateNumber(num, cell, styles, dateCache))
      {
        // Only the calendar fields are used, so the CTRL time range is irrelevant.
        std::tm tm{};
        int16_t milli = 0;
        time_t sec = 0;
        serialToLocalTime(num.value, tm, milli, sec, std::numeric_limits<time_t>::max(), state);

        bool hasTime = tm.tm_hour != 0 || tm.tm_min != 0 || tm.tm_sec != 0;
        char buf[32];
        strftime(buf, sizeof(buf), hasTime ? "%Y-%m-%d %H:%M:%S" : "%Y-%m-%d", &tm);
        return buf;
      }

      if (num.isInteger)
        return std::to_string(num.integer);

      double intpart;
      if (modf(num.value, &intpart) == 0.0 && fabs(intpart) < 1e15)
        return std::to_string(static_cast<long long>(intpart));
      char buf[32];
      snprintf(buf, sizeof(buf), "%.15g", num.value);
      return buf;
    }

    switch (val.type())
    {
      case XLValueType::String:
        return decodeExcelEscapes(val.get<std::string>());
      case XLValueType::Boolean:
        return val.get<bool>() ? "TRUE" : "FALSE";
      default:
        return std::string();
    }
  }

  //----------------------------------------------------------------------------
  // Sheet names
  //----------------------------------------------------------------------------

  std::string checkSheetNames(std::vector<std::string> &names)
  {
    std::unordered_set<std::u32string> seen;
    for (std::string &name : names)
    {
      name = toXmlUtf8(name.c_str(), false);

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

      // Excel compares sheet names case-insensitively.
      std::u32string folded = foldCase(name);
      if (folded == U"history")
        return quoted + " is reserved by Excel";
      if (!seen.insert(folded).second)
        return quoted + " duplicates another sheet name (Excel ignores case)";
    }
    return std::string();
  }
}
