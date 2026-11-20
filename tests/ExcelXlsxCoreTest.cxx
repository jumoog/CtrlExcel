// Tests of ExcelXlsxCore and, through it, the xlsx library (no WinCC OA API
// needed). Run through CTest, or directly; an optional argument replaces the
// path of testdata/CtrlExcelReaderFixture.xlsx.

#include <ExcelXlsxCore.hxx>
#include <FixturePath.hxx>

#include <chrono>
#include <cmath>
#include <cstdio>
#include <cstdlib>
#include <cstring>
#include <ctime>
#include <exception>
#include <filesystem>
#include <iostream>
#include <limits>
#include <sstream>
#include <stdexcept>
#include <string>
#include <vector>

using namespace ExcelXlsxCore;

namespace
{
  int failures = 0;
  std::string fixturePath = CTRLEXCEL_FIXTURE; // UTF-8

  void report(bool ok, const std::string &what, const char *file, int line)
  {
    if (ok)
      return;
    ++failures;
    std::cerr << file << "(" << line << "): FAILED " << what << "\n";
  }

  template <typename A, typename B>
  void checkEqual(const A &actual, const B &expected, const char *expr, const char *file, int line)
  {
    if (actual == expected)
      return;
    std::ostringstream msg;
    msg << expr << ": got <" << actual << ">, expected <" << expected << ">";
    report(false, msg.str(), file, line);
  }

  // Sets the time zone of mktime/localtime: POSIX TZ syntax, which both glibc
  // and MSVC accept for fixed offsets (MSVC ignores DST rules, see tzTests).
  void setTimeZone(const char *tz)
  {
#ifdef _WIN32
    _putenv_s("TZ", tz);
    _tzset();
#else
    setenv("TZ", tz, 1);
    tzset();
#endif
  }

  // A temporary .xlsx path, removed when the object goes out of scope.
  struct TempFile
  {
    std::filesystem::path path;

    explicit TempFile(const char *name)
    {
      long long stamp = std::chrono::steady_clock::now().time_since_epoch().count();
      path = std::filesystem::temp_directory_path()
           / ("ExcelXlsxCoreTest_" + std::to_string(stamp) + "_" + name + ".xlsx");
    }
    ~TempFile()
    {
      std::error_code ec;
      std::filesystem::remove(path, ec);
    }
    std::string name() const { return path.u8string(); }
  };

}

#define CHECK(cond) report((cond), #cond, __FILE__, __LINE__)
#define CHECK_EQ(actual, expected) checkEqual((actual), (expected), #actual, __FILE__, __LINE__)

//------------------------------------------------------------------------------
// Calendar arithmetic
//------------------------------------------------------------------------------

void calendarTest()
{
  CHECK_EQ(daysFromCivil(1970, 1, 1), 0LL);
  CHECK_EQ(daysFromCivil(1900, 1, 1), -25567LL);
  CHECK_EQ(daysFromCivil(2000, 3, 1), 11017LL);
  CHECK_EQ(daysFromCivil(2026, 1, 1), 20454LL);

  bool roundTrips = true;
  for (long long z = -800000; z <= 800000 && roundTrips; ++z)
  {
    long long y;
    unsigned m, d;
    civilFromDays(z, y, m, d);
    roundTrips = daysFromCivil(y, m, d) == z;
  }
  CHECK(roundTrips);

  // Excel serial days, including the fictitious 1900-02-29 (serial 60).
  CHECK_EQ(excelDayToUnixDay(1), daysFromCivil(1900, 1, 1));
  CHECK_EQ(excelDayToUnixDay(59), daysFromCivil(1900, 2, 28));
  CHECK_EQ(excelDayToUnixDay(61), daysFromCivil(1900, 3, 1));
  CHECK_EQ(excelDayToUnixDay(25569), 0LL);
  CHECK_EQ(excelDayToUnixDay(46023), daysFromCivil(2026, 1, 1));

  bool inverse = true;
  for (long long excelDay = 1; excelDay <= 200000 && inverse; ++excelDay)
    if (excelDay != 60)
      inverse = unixDayToExcelDay(excelDayToUnixDay(excelDay)) == excelDay;
  CHECK(inverse);
}

//------------------------------------------------------------------------------
// Text
//------------------------------------------------------------------------------

void utf16LengthTest()
{
  CHECK_EQ(utf16Length(""), size_t(0));
  CHECK_EQ(utf16Length("abc"), size_t(3));
  CHECK_EQ(utf16Length(u8"ä€"), size_t(2));   // ä€
  CHECK_EQ(utf16Length(u8"\U0001F600"), size_t(2));      // surrogate pair
}

void toValidUtf8Test()
{
  CHECK_EQ(toValidUtf8("plain"), std::string("plain"));
  // Nothing is escaped: the xlsx library escapes what XML cannot hold.
  CHECK_EQ(toValidUtf8("a\rb\x01"), std::string("a\rb\x01"));
  CHECK_EQ(toValidUtf8("AB_x0041_7"), std::string("AB_x0041_7"));
  CHECK_EQ(toValidUtf8(u8"\u00e4\U0001F600"), std::string(u8"\u00e4\U0001F600"));

  // Invalid UTF-8 bytes are taken as ISO-8859-1.
  CHECK_EQ(toValidUtf8("\xE4"), std::string(u8"\u00e4"));     // Latin-1 a-umlaut
  CHECK_EQ(toValidUtf8("\xC0\xAF"), std::string("\xC3\x80\xC2\xAF")); // overlong '/'
  CHECK_EQ(toValidUtf8("\xED\xA0\x80"), std::string("\xC3\xAD\xC2\xA0\xC2\x80")); // surrogate
  CHECK_EQ(toValidUtf8("\xF4\x90\x80\x80"), std::string("\xC3\xB4\xC2\x90\xC2\x80\xC2\x80")); // > U+10FFFF
  CHECK_EQ(toValidUtf8("\xC3"), std::string("\xC3\x83"));      // truncated sequence
}

void foldCaseTest()
{
  CHECK(foldCase("Sheet") == foldCase("sHEET"));
  CHECK(foldCase(u8"Äpfel") == foldCase(u8"äpfel"));   // Ä ä
  CHECK(foldCase(u8"Σ") == foldCase(u8"σ"));           // Σ σ
  CHECK(foldCase(u8"ЁЖ") == foldCase(u8"ёж")); // ЁЖ ёж
  CHECK(foldCase(u8"ŁĀ") == foldCase(u8"łā")); // Łā
  CHECK(foldCase(u8"×") != foldCase(u8"÷"));           // × ÷ are no pair
  CHECK(foldCase("a") != foldCase("b"));
  CHECK(foldCase(u8"\U0001F600") == std::u32string(1, U'\U0001F600'));
}

void checkSheetNamesTest()
{
  auto problem = [](std::vector<std::string> names) { return checkSheetNames(names); };

  CHECK_EQ(problem({"Data", "Summary", u8"Äpfel"}), std::string());
  CHECK_EQ(problem({std::string(31, 'x')}), std::string());
  CHECK(!problem({std::string(32, 'x')}).empty());
  // 15 surrogate pairs + 1 = 31 UTF-16 units; one more pair exceeds the limit.
  std::string emoji;
  for (int i = 0; i < 15; ++i)
    emoji += u8"\U0001F600";
  CHECK_EQ(problem({emoji + "a"}), std::string());
  CHECK(!problem({emoji + u8"\U0001F600"}).empty());

  CHECK(!problem({""}).empty());
  for (const char *bad : {"a/b", "a\\b", "a?b", "a*b", "a[b", "a]b", "a:b", "a\tb", "'a", "a'"})
    CHECK(!problem({bad}).empty());
  CHECK_EQ(problem({"a'b"}), std::string());

  CHECK(!problem({"History"}).empty());
  CHECK(!problem({"Data", "DATA"}).empty());
  CHECK(!problem({u8"Äpfel", u8"äPFEL"}).empty());

  // Names are converted to valid UTF-8 in place.
  std::vector<std::string> names = {"\xE4"};
  CHECK_EQ(checkSheetNames(names), std::string());
  CHECK_EQ(names[0], std::string(u8"ä"));
}

void numberWidthTest()
{
  CHECK_EQ(numberWidth(12345), size_t(5));
  CHECK_EQ(numberWidth(-1.5), size_t(4));
  CHECK_EQ(numberWidth(0.1), size_t(3));
  CHECK_EQ(numberWidth(1e20), size_t(5)); // "1e+20"
}

//------------------------------------------------------------------------------
// Date formats and conversion
//------------------------------------------------------------------------------

void utcTimeTest()
{
  setTimeZone("UTC0");
  ReadState state;
  const time_t maxSec = std::numeric_limits<time_t>::max();

  std::tm tm{};
  int16_t milli = -1;
  time_t sec = -1;

  CHECK(serialToLocalTime(25569.0, tm, milli, sec, maxSec, state) == LocalTime::Ok);
  CHECK_EQ(static_cast<long long>(sec), 0LL);
  CHECK_EQ(milli, int16_t(0));

  // 2026-01-01 12:00:00.999: milliseconds are part of the serial.
  double serial = 46023.0 + 43200999.0 / static_cast<double>(DAY_MILLIS);
  CHECK(serialToLocalTime(serial, tm, milli, sec, maxSec, state) == LocalTime::Ok);
  CHECK_EQ(static_cast<long long>(sec), 20454LL * DAY_SECONDS + 43200);
  CHECK_EQ(milli, int16_t(999));
  CHECK_EQ(tm.tm_year, 126);
  CHECK_EQ(tm.tm_hour, 12);

  // Rounding to milliseconds must not leave 23:59:59.1000.
  CHECK(serialToLocalTime(46023.0 - 1e-12, tm, milli, sec, maxSec, state) == LocalTime::Ok);
  CHECK_EQ(static_cast<long long>(sec), 20454LL * DAY_SECONDS);
  CHECK_EQ(milli, int16_t(0));

  // Outside CTRL time: before 1970 (calendar fields still set) or past maxSec.
  CHECK(serialToLocalTime(22037.0, tm, milli, sec, maxSec, state) == LocalTime::OutOfRange);
  CHECK_EQ(tm.tm_year, 60);
  CHECK_EQ(tm.tm_mon, 4);
  CHECK_EQ(tm.tm_mday, 1);
  CHECK(serialToLocalTime(46023.0, tm, milli, sec, 1000, state) == LocalTime::OutOfRange);

  double out = 0.0;
  CHECK(localTimeToSerial(0, 0, out));
  CHECK_EQ(out, 25569.0);
  CHECK(localTimeToSerial(static_cast<time_t>(20454LL * DAY_SECONDS + 43200), 999, out));
  CHECK(std::fabs(out - serial) < 1e-9);

  // Round trip over 1970..2100 (every ~37 hours, with milliseconds).
  bool roundTrips = true;
  for (long long t = 0; t < 4102444800LL && roundTrips; t += 133333)
  {
    int ms = static_cast<int>(t % 1000);
    roundTrips = localTimeToSerial(static_cast<time_t>(t), ms, out)
              && serialToLocalTime(out, tm, milli, sec, maxSec, state) == LocalTime::Ok
              && static_cast<long long>(sec) == t && milli == ms;
  }
  CHECK(roundTrips);
}

#ifndef _WIN32
// MSVC applies US DST rules to any TZ value, so this runs on POSIX only.
void dstTimeTest()
{
  setTimeZone("CET-1CEST,M3.5.0,M10.5.0/3");
  ReadState state;
  const time_t maxSec = std::numeric_limits<time_t>::max();

  std::tm tm{};
  int16_t milli = 0;
  time_t sec = 0;

  // 2026-01-01 00:00 CET is 2025-12-31 23:00 UTC.
  CHECK(serialToLocalTime(46023.0, tm, milli, sec, maxSec, state) == LocalTime::Ok);
  CHECK_EQ(static_cast<long long>(sec), 20454LL * DAY_SECONDS - 3600);

  // 2026-07-01 12:00 CEST is 10:00 UTC.
  long long july = daysFromCivil(2026, 7, 1);
  CHECK(serialToLocalTime(static_cast<double>(unixDayToExcelDay(july)) + 0.5, tm, milli, sec, maxSec, state)
        == LocalTime::Ok);
  CHECK_EQ(static_cast<long long>(sec), july * DAY_SECONDS + 10 * 3600);

  // 2026-03-29 02:30 does not exist (clocks jump from 02:00 to 03:00).
  long long spring = daysFromCivil(2026, 3, 29);
  double skipped = static_cast<double>(unixDayToExcelDay(spring)) + 2.5 / 24.0;
  CHECK(serialToLocalTime(skipped, tm, milli, sec, maxSec, state) == LocalTime::Shifted);
  // 01:30 and 03:30 on that day exist.
  CHECK(serialToLocalTime(skipped - 1.0 / 24.0, tm, milli, sec, maxSec, state) == LocalTime::Ok);
  CHECK_EQ(static_cast<long long>(sec), spring * DAY_SECONDS + 30 * 60);
  CHECK(serialToLocalTime(skipped + 1.0 / 24.0, tm, milli, sec, maxSec, state) == LocalTime::Ok);
  CHECK_EQ(static_cast<long long>(sec), spring * DAY_SECONDS + 90 * 60);

  // Hourly round trip through 2026, except the repeated hour after the
  // fall-back, whose wall-clock times are ambiguous.
  double out = 0.0, before = 0.0, after = 0.0;
  bool roundTrips = true;
  long long start = daysFromCivil(2026, 1, 1) * DAY_SECONDS;
  for (long long t = start; t < start + 366 * DAY_SECONDS && roundTrips; t += 1800)
  {
    localTimeToSerial(static_cast<time_t>(t - 3600), 0, before);
    localTimeToSerial(static_cast<time_t>(t + 3600), 0, after);
    roundTrips = localTimeToSerial(static_cast<time_t>(t), 0, out);
    if (out == before || out == after)
      continue;
    roundTrips = serialToLocalTime(out, tm, milli, sec, maxSec, state) == LocalTime::Ok
              && static_cast<long long>(sec) == t;
  }
  CHECK(roundTrips);
  setTimeZone("UTC0");
}
#endif

//------------------------------------------------------------------------------
// Cell values
//------------------------------------------------------------------------------

void integerFitTest()
{
  int64_t i = 0;
  CHECK(integerFit(87.0, i) == IntegerFit::Int);
  CHECK_EQ(i, int64_t(87));
  CHECK(integerFit(-0.0, i) == IntegerFit::Int);
  CHECK_EQ(i, int64_t(0));
  CHECK(integerFit(2147483647.0, i) == IntegerFit::Int);
  CHECK(integerFit(-2147483648.0, i) == IntegerFit::Int);
  CHECK(integerFit(2147483648.0, i) == IntegerFit::Long);
  CHECK(integerFit(5000000000.0, i) == IntegerFit::Long);
  CHECK_EQ(i, int64_t(5000000000LL));
  CHECK(integerFit(-9223372036854775808.0, i) == IntegerFit::Long);
  CHECK_EQ(i, std::numeric_limits<int64_t>::min());
  CHECK(integerFit(9223372036854775808.0, i) == IntegerFit::None); // 2^63
  CHECK(integerFit(1e20, i) == IntegerFit::None);
  CHECK(integerFit(0.25, i) == IntegerFit::None);
  CHECK(integerFit(std::nan(""), i) == IntegerFit::None);
  CHECK(integerFit(std::numeric_limits<double>::infinity(), i) == IntegerFit::None);
}

cxlsx_cell makeCell(uint32_t type, double number = 0.0, const char *text = nullptr, bool boolean = false)
{
  cxlsx_cell cell{};
  cell.row = 1;
  cell.column = 1;
  cell.type = type;
  cell.number = number;
  cell.boolean = boolean;
  cell.text = text;
  cell.text_len = text ? std::strlen(text) : 0;
  return cell;
}

void headerTextTest()
{
  setTimeZone("UTC0");
  ReadState state;
  CHECK_EQ(headerText(makeCell(CXLSX_TEXT, 0, "Name"), state), std::string("Name"));
  CHECK_EQ(headerText(makeCell(CXLSX_NUMBER, 2024), state), std::string("2024"));
  CHECK_EQ(headerText(makeCell(CXLSX_NUMBER, 2.5), state), std::string("2.5"));
  CHECK_EQ(headerText(makeCell(CXLSX_NUMBER, 1e20), state), std::string("1e+20"));
  CHECK_EQ(headerText(makeCell(CXLSX_BOOL, 0, nullptr, true), state), std::string("TRUE"));
  CHECK_EQ(headerText(makeCell(CXLSX_DATETIME, 46023.5), state), std::string("2026-01-01 12:00:00"));
  CHECK_EQ(headerText(makeCell(CXLSX_DATETIME, 46023.0), state), std::string("2026-01-01"));
  CHECK_EQ(headerText(makeCell(CXLSX_DATETIME, 22037.0), state), std::string("1960-05-01"));
  // A time-only serial and a duration stay numbers.
  CHECK_EQ(headerText(makeCell(CXLSX_DATETIME, 0.5), state), std::string("0.5"));
  CHECK_EQ(headerText(makeCell(CXLSX_DURATION, 1.5), state), std::string("1.5"));
  CHECK_EQ(headerText(makeCell(CXLSX_ERROR, 0, "#N/A"), state), std::string());
  CHECK_EQ(headerText(makeCell(CXLSX_UNCACHED_FORMULA), state), std::string());

  CHECK(isDateCell(makeCell(CXLSX_DATETIME, 1.0)));
  CHECK(!isDateCell(makeCell(CXLSX_DATETIME, 0.99)));
  CHECK(!isDateCell(makeCell(CXLSX_NUMBER, 46023.0)));
}

//------------------------------------------------------------------------------
// The xlsx library
//------------------------------------------------------------------------------

namespace
{
  struct RecordedCell
  {
    uint32_t row;
    uint32_t column;
    uint32_t type;
    double number;
    bool boolean;
    std::string text;
  };

  // Everything Workbook::readSheet delivers; optionally throws at a cell.
  struct Recorder : SheetVisitor
  {
    uint32_t rows = 0, columns = 0;
    std::vector<std::pair<uint32_t, bool>> rowList;
    std::vector<RecordedCell> cells;
    bool rowBeforeCells = true;
    size_t throwAt = SIZE_MAX;

    bool begin(uint32_t r, uint32_t c) override
    {
      rows = r;
      columns = c;
      return true;
    }

    bool row(uint32_t number, bool hidden) override
    {
      rowList.emplace_back(number, hidden);
      return true;
    }

    bool cell(const cxlsx_cell &c) override
    {
      if (cells.size() == throwAt)
        throw std::logic_error("thrown by the visitor");
      rowBeforeCells = rowBeforeCells && !rowList.empty() && rowList.back().first == c.row;
      cells.push_back({c.row, c.column, c.type, c.number, c.boolean, cellText(c)});
      return true;
    }

    const RecordedCell *at(uint32_t row, uint32_t column) const
    {
      for (const RecordedCell &c : cells)
        if (c.row == row && c.column == column)
          return &c;
      return nullptr;
    }
  };

  // The cell at (row, column) has this type (and number or text).
  bool cellIs(const Recorder &rec, uint32_t row, uint32_t column, uint32_t type,
              double number = 0.0, const std::string &text = std::string())
  {
    const RecordedCell *c = rec.at(row, column);
    if (!c || c->type != type)
      return false;
    if (type == CXLSX_TEXT || type == CXLSX_ERROR)
      return c->text == text;
    if (type == CXLSX_BOOL)
      return c->boolean == (number != 0.0);
    return type == CXLSX_UNCACHED_FORMULA || c->number == number;
  }
}

// testdata/CtrlExcelReaderFixture.xlsx (written like Excel does, see
// testdata/make_fixture.py) through the library.
void fixtureTest()
{
  if (!std::filesystem::exists(std::filesystem::u8path(fixturePath)))
  {
    report(false, "fixture not found: '" + fixturePath + "'", __FILE__, __LINE__);
    return;
  }

  Workbook book(fixturePath);
  CHECK(book.sheetNames() == std::vector<std::string>({"Fixture", "Implicit"}));

  Recorder rec;
  book.readSheet("Fixture", rec);
  CHECK_EQ(rec.rows, uint32_t(12));
  CHECK_EQ(rec.columns, uint32_t(2));
  CHECK_EQ(rec.rowList.size(), size_t(12)); // row 11 only carries formatting
  CHECK(rec.rowBeforeCells);
  CHECK(cellIs(rec, 2, 2, CXLSX_DATETIME, 22037.0));  // pre1970
  CHECK(cellIs(rec, 3, 2, CXLSX_DATETIME, 46023.0));  // dateInt: whole serial, still a date
  CHECK(cellIs(rec, 4, 2, CXLSX_DURATION, 1.5));      // [h]:mm
  CHECK(cellIs(rec, 5, 2, CXLSX_ERROR, 0, "#DIV/0!"));
  CHECK(cellIs(rec, 6, 2, CXLSX_UNCACHED_FORMULA));
  CHECK(cellIs(rec, 7, 2, CXLSX_TEXT, 0, "a\r\nb"));  // _x000D_
  CHECK(cellIs(rec, 8, 2, CXLSX_TEXT, 0, "ABA7"));    // every escape is decoded
  CHECK(cellIs(rec, 9, 2, CXLSX_NUMBER, 1e20));       // "1E+20"
  CHECK(cellIs(rec, 10, 2, CXLSX_NUMBER, 12345678901234567890.0));
  CHECK(rec.at(11, 1) == nullptr);
  CHECK(cellIs(rec, 12, 2, CXLSX_TEXT, 0, "last"));

  // Sheet "Implicit": row 3 missing, rows 5 and 6 and some cells without r,
  // row 5 ends in <extLst>.
  Recorder implicit;
  book.readSheet("Implicit", implicit);
  CHECK_EQ(implicit.rows, uint32_t(6));
  CHECK_EQ(implicit.columns, uint32_t(2));
  std::vector<uint32_t> rowNumbers;
  for (const auto &r : implicit.rowList)
    rowNumbers.push_back(r.first);
  CHECK(rowNumbers == std::vector<uint32_t>({1, 2, 4, 5, 6}));
  CHECK(cellIs(implicit, 5, 1, CXLSX_TEXT, 0, "implicitRow"));
  CHECK(cellIs(implicit, 5, 2, CXLSX_NUMBER, 3.0));
  CHECK(cellIs(implicit, 6, 2, CXLSX_NUMBER, 4.0));
}

// WorkbookWriter, read back through Workbook.
void writeReadTest()
{
  TempFile file("write_read");
  {
    WorkbookWriter out;
    out.addSheet("Data");
    out.text(1, 1, "Name", CellStyle::Header);
    out.text(1, 2, "", CellStyle::Header); // empty header: a styled blank cell
    out.text(2, 1, "a\rb\x01_x0041_\xEF\xBF\xBE");
    out.number(2, 2, 46023.5, CellStyle::DateTime);
    out.boolean(2, 3, true);
    out.number(3, 1, 5000000000.0);
    out.number(3, 2, std::numeric_limits<double>::infinity());
    out.number(3, 3, -3e-7);
    out.text(3, 4, ""); // not stored
    out.columnWidth(1, 1000.0); // capped
    out.addSheet(u8"Äpfel");
    out.number(1, 1, 1.0);
    out.save(file.name());
  }

  Workbook book(file.name());
  CHECK(book.sheetNames() == std::vector<std::string>({"Data", u8"Äpfel"}));
  Recorder rec;
  book.readSheet("Data", rec);
  CHECK_EQ(rec.rows, uint32_t(3));
  CHECK_EQ(rec.columns, uint32_t(3)); // the blank header cell counts, D3 was not written
  CHECK(cellIs(rec, 1, 1, CXLSX_TEXT, 0, "Name"));
  CHECK(rec.at(1, 2) == nullptr);
  CHECK(cellIs(rec, 2, 1, CXLSX_TEXT, 0, "a\rb\x01_x0041_\xEF\xBF\xBE"));
  CHECK(cellIs(rec, 2, 2, CXLSX_DATETIME, 46023.5));
  CHECK(cellIs(rec, 2, 3, CXLSX_BOOL, 1.0));
  CHECK(cellIs(rec, 3, 1, CXLSX_NUMBER, 5000000000.0));
  CHECK(cellIs(rec, 3, 2, CXLSX_ERROR, 0, "#NUM!"));
  CHECK(cellIs(rec, 3, 3, CXLSX_NUMBER, -3e-7));

  Recorder second;
  book.readSheet(u8"Äpfel", second);
  CHECK(cellIs(second, 1, 1, CXLSX_NUMBER, 1.0));

  // A workbook without sheets gets "Sheet1"; save replaces the file.
  {
    WorkbookWriter empty;
    empty.save(file.name());
  }
  CHECK(Workbook(file.name()).sheetNames() == std::vector<std::string>({"Sheet1"}));
}

// Errors of the library and of visitors arrive as exceptions.
void errorTest()
{
  bool thrown = false;
  try
  {
    Workbook missing("does/not/exist.xlsx");
  }
  catch (const std::runtime_error &e)
  {
    thrown = std::string(e.what()).size() > 0;
  }
  CHECK(thrown);

  Workbook book(fixturePath);
  Recorder rec;
  thrown = false;
  try
  {
    book.readSheet("No such sheet", rec);
  }
  catch (const std::runtime_error &e)
  {
    thrown = std::string(e.what()).find("No such sheet") != std::string::npos;
  }
  CHECK(thrown);

  // An exception of the visitor ends the read and is rethrown as it is.
  Recorder throwing;
  throwing.throwAt = 2;
  thrown = false;
  try
  {
    book.readSheet("Fixture", throwing);
  }
  catch (const std::logic_error &)
  {
    thrown = true;
  }
  CHECK(thrown);
  CHECK_EQ(throwing.cells.size(), size_t(2));
  // The book stays usable.
  Recorder after;
  book.readSheet("Implicit", after);
  CHECK_EQ(after.rowList.size(), size_t(5));

  WorkbookWriter out;
  thrown = false;
  try
  {
    out.text(1, 1, "no sheet added yet");
  }
  catch (const std::runtime_error &)
  {
    thrown = true;
  }
  CHECK(thrown);
  out.addSheet("S");
  thrown = false;
  try
  {
    out.text(1, 1, std::string(32768, 'a')); // over Excel's cell limit
  }
  catch (const std::runtime_error &)
  {
    thrown = true;
  }
  CHECK(thrown);

  // A directory of the target name cannot be replaced: save fails and
  // leaves it alone.
  TempFile dir("directory");
  std::filesystem::create_directory(dir.path);
  thrown = false;
  try
  {
    out.save(dir.name());
  }
  catch (const std::runtime_error &)
  {
    thrown = true;
  }
  CHECK(thrown);
  CHECK(std::filesystem::is_directory(dir.path));
}

//------------------------------------------------------------------------------

int main(int argc, char *argv[])
{
  // path(char *) decodes the native narrow encoding (the ANSI code page on
  // Windows), so non-ASCII paths survive the conversion to UTF-8.
  if (argc > 1)
    fixturePath = std::filesystem::path(argv[1]).u8string();

  const std::pair<const char *, void (*)()> tests[] = {
    {"calendarTest", calendarTest},
    {"utf16LengthTest", utf16LengthTest},
    {"toValidUtf8Test", toValidUtf8Test},
    {"foldCaseTest", foldCaseTest},
    {"checkSheetNamesTest", checkSheetNamesTest},
    {"numberWidthTest", numberWidthTest},
    {"integerFitTest", integerFitTest},
    {"utcTimeTest", utcTimeTest},
#ifndef _WIN32
    {"dstTimeTest", dstTimeTest},
#endif
    {"headerTextTest", headerTextTest},
    {"fixtureTest", fixtureTest},
    {"writeReadTest", writeReadTest},
    {"errorTest", errorTest},
  };

  const int numTests = static_cast<int>(sizeof(tests) / sizeof(tests[0]));
  int failedTests = 0;
  for (const auto &test : tests)
  {
    int before = failures;
    try
    {
      test.second();
    }
    catch (const std::exception &e)
    {
      report(false, std::string("exception: ") + e.what(), __FILE__, __LINE__);
    }
    bool pass = failures == before;
    failedTests += pass ? 0 : 1;
    std::cout << (pass ? "pass " : "FAIL ") << test.first << "\n";
  }

  std::cout << (failedTests == 0 ? "ALL PASSED" : "FAILURES") << ": "
            << (numTests - failedTests) << " passed, "
            << failedTests << " failed\n";
  return failedTests == 0 ? EXIT_SUCCESS : EXIT_FAILURE;
}
