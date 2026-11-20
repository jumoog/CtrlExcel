#include <ExcelXlsxCore.hxx>

#include <cmath>
#include <cstdio>
#include <cstring>
#include <exception>
#include <limits>
#include <new>
#include <stdexcept>
#include <unordered_set>

namespace
{
  constexpr size_t EXCEL_MAX_SHEET_NAME_CHARS = 31;

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

  std::string toValidUtf8(const char *text)
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

      out.append(reinterpret_cast<const char *>(p), len);
      p += len;
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
  // Sheet names
  //----------------------------------------------------------------------------

  std::string checkSheetNames(std::vector<std::string> &names)
  {
    std::unordered_set<std::u32string> seen;
    for (std::string &name : names)
    {
      name = toValidUtf8(name.c_str());

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

  //----------------------------------------------------------------------------
  // Cell values
  //----------------------------------------------------------------------------

  IntegerFit integerFit(double value, int64_t &integer)
  {
    double intpart;
    if (!std::isfinite(value) || std::modf(value, &intpart) != 0.0)
      return IntegerFit::None;
    // 2^63 is exactly representable; every whole double below it converts.
    if (intpart < -9223372036854775808.0 || intpart >= 9223372036854775808.0)
      return IntegerFit::None;
    integer = static_cast<int64_t>(intpart);
    return (integer >= INT_MIN && integer <= INT_MAX) ? IntegerFit::Int : IntegerFit::Long;
  }

  bool isDateCell(const cxlsx_cell &cell)
  {
    return cell.type == CXLSX_DATETIME && cell.number >= 1.0;
  }

  std::string headerText(const cxlsx_cell &cell, ReadState &state)
  {
    switch (cell.type)
    {
      case CXLSX_TEXT:
        return cellText(cell);
      case CXLSX_BOOL:
        return cell.boolean ? "TRUE" : "FALSE";
      case CXLSX_DATETIME:
      case CXLSX_DURATION:
      case CXLSX_NUMBER:
        break;
      default: // errors, uncalculated formulas
        return std::string();
    }

    if (isDateCell(cell))
    {
      // Only the calendar fields are used, so the CTRL time range is irrelevant.
      std::tm tm{};
      int16_t milli = 0;
      time_t sec = 0;
      serialToLocalTime(cell.number, tm, milli, sec, std::numeric_limits<time_t>::max(), state);

      bool hasTime = tm.tm_hour != 0 || tm.tm_min != 0 || tm.tm_sec != 0;
      char buf[32];
      strftime(buf, sizeof(buf), hasTime ? "%Y-%m-%d %H:%M:%S" : "%Y-%m-%d", &tm);
      return buf;
    }

    // Whole numbers without decimals, as long as a double holds them exactly.
    int64_t integer = 0;
    if (integerFit(cell.number, integer) != IntegerFit::None && std::fabs(cell.number) < 9007199254740992.0)
      return std::to_string(integer);
    char buf[32];
    snprintf(buf, sizeof(buf), "%.15g", cell.number);
    return buf;
  }

  //----------------------------------------------------------------------------
  // The xlsx library
  //----------------------------------------------------------------------------

  namespace
  {
    // Receives the library's error message. set() runs inside the library
    // and must not throw.
    struct ErrorSink
    {
      std::string message = "unknown error";
      cxlsx_error c{this, &ErrorSink::set};

      static void set(void *ctx, const char *msg, size_t len)
      {
        try
        {
          static_cast<ErrorSink *>(ctx)->message.assign(msg, len);
        }
        catch (...)
        {
        }
      }

      [[noreturn]] void raise() const { throw std::runtime_error(message); }
    };

    void check(bool ok, const ErrorSink &error)
    {
      if (!ok)
        error.raise();
    }

    // Calls the SheetVisitor from the library's callbacks; an exception is
    // kept and rethrown after the library returns, it must not unwind
    // through Rust.
    struct VisitorThunk
    {
      SheetVisitor &visitor;
      std::exception_ptr error;

      template <typename Fn>
      static bool call(void *ctx, Fn &&fn)
      {
        VisitorThunk &self = *static_cast<VisitorThunk *>(ctx);
        try
        {
          return fn(self.visitor);
        }
        catch (...)
        {
          self.error = std::current_exception();
          return false;
        }
      }

      static bool begin(void *ctx, uint32_t rows, uint32_t columns)
      {
        return call(ctx, [&](SheetVisitor &v) { return v.begin(rows, columns); });
      }

      static bool row(void *ctx, uint32_t row, bool hidden)
      {
        return call(ctx, [&](SheetVisitor &v) { return v.row(row, hidden); });
      }

      static bool cell(void *ctx, const cxlsx_cell *cell)
      {
        return call(ctx, [&](SheetVisitor &v) { return v.cell(*cell); });
      }
    };
  }

  Workbook::Workbook(const std::string &path)
  {
    ErrorSink error;
    m_book = cxlsx_book_open(path.data(), path.size(), &error.c);
    if (!m_book)
      error.raise();
  }

  Workbook::~Workbook()
  {
    cxlsx_book_close(m_book);
  }

  std::vector<std::string> Workbook::sheetNames() const
  {
    std::vector<std::string> names;
    const size_t count = cxlsx_book_sheet_count(m_book);
    names.reserve(count);
    for (size_t i = 0; i < count; ++i)
    {
      size_t len = 0;
      const char *name = cxlsx_book_sheet_name(m_book, i, &len);
      names.emplace_back(name ? name : "", len);
    }
    return names;
  }

  void Workbook::readSheet(const std::string &name, SheetVisitor &visitor)
  {
    VisitorThunk thunk{visitor, nullptr};
    const cxlsx_sheet_visitor callbacks{&thunk, &VisitorThunk::begin, &VisitorThunk::row, &VisitorThunk::cell};
    ErrorSink error;
    int status = cxlsx_book_read_sheet(m_book, name.data(), name.size(), &callbacks, &error.c);
    if (thunk.error)
      std::rethrow_exception(thunk.error);
    if (status == CXLSX_FAILED)
      error.raise();
  }

  WorkbookWriter::WorkbookWriter()
    : m_writer(cxlsx_writer_new())
  {
    if (!m_writer)
      throw std::bad_alloc();
  }

  WorkbookWriter::~WorkbookWriter()
  {
    cxlsx_writer_free(m_writer);
  }

  void WorkbookWriter::addSheet(const std::string &name)
  {
    ErrorSink error;
    check(cxlsx_writer_add_sheet(m_writer, name.data(), name.size(), &error.c), error);
  }

  void WorkbookWriter::text(uint32_t row, uint32_t column, const std::string &utf8, CellStyle style)
  {
    ErrorSink error;
    check(cxlsx_writer_text(m_writer, row, column, utf8.data(), utf8.size(),
                            static_cast<uint32_t>(style), &error.c), error);
  }

  void WorkbookWriter::number(uint32_t row, uint32_t column, double value, CellStyle style)
  {
    ErrorSink error;
    check(cxlsx_writer_number(m_writer, row, column, value, static_cast<uint32_t>(style), &error.c), error);
  }

  void WorkbookWriter::boolean(uint32_t row, uint32_t column, bool value, CellStyle style)
  {
    ErrorSink error;
    check(cxlsx_writer_bool(m_writer, row, column, value, static_cast<uint32_t>(style), &error.c), error);
  }

  void WorkbookWriter::columnWidth(uint32_t column, double width)
  {
    ErrorSink error;
    check(cxlsx_writer_column_width(m_writer, column, width, &error.c), error);
  }

  void WorkbookWriter::save(const std::string &path)
  {
    ErrorSink error;
    check(cxlsx_writer_save(m_writer, path.data(), path.size(), &error.c), error);
  }
}
