// Tests of ExcelXlsxCore (no WinCC OA API needed). Run through CTest, or
// directly; an optional argument replaces the path of
// testdata/CtrlExcelReaderFixture.xlsx.

#include <ExcelXlsxCore.hxx>
#include <FixturePath.hxx>

#include <chrono>
#include <cmath>
#include <cstdio>
#include <cstdlib>
#include <ctime>
#include <exception>
#include <filesystem>
#include <iostream>
#include <limits>
#include <sstream>
#include <stdexcept>
#include <string>
#include <vector>

using namespace OpenXLSX;
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

  // cellNumber() of the cell at ref.
  bool numberAt(XLWorksheet &wks, const char *ref, CellNumber &num)
  {
    XLCell cell = wks.cell(ref);
    return cellNumber(cell, cell.value(), num);
  }
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

void toXmlUtf8Test()
{
  CHECK_EQ(toXmlUtf8("plain"), std::string("plain"));
  CHECK_EQ(toXmlUtf8("tab\tlf\n"), std::string("tab\tlf\n"));
  CHECK_EQ(toXmlUtf8("a\rb"), std::string("a_x000D_b"));
  CHECK_EQ(toXmlUtf8("\x01" "x"), std::string("_x0001_x"));
  CHECK_EQ(toXmlUtf8("\xEF\xBF\xBE\xEF\xBF\xBF"), std::string("_xFFFE__xFFFF_"));
  CHECK_EQ(toXmlUtf8("AB_x0041_7"), std::string("AB_x005F_x0041_7"));
  CHECK_EQ(toXmlUtf8(u8"ä\U0001F600"), std::string(u8"ä\U0001F600"));

  // Invalid UTF-8 bytes are taken as ISO-8859-1.
  CHECK_EQ(toXmlUtf8("\xE4"), std::string(u8"ä"));        // Latin-1 ä
  CHECK_EQ(toXmlUtf8("\xC0\xAF"), std::string("\xC3\x80\xC2\xAF")); // overlong '/'
  CHECK_EQ(toXmlUtf8("\xED\xA0\x80"), std::string("\xC3\xAD\xC2\xA0\xC2\x80")); // surrogate
  CHECK_EQ(toXmlUtf8("\xF4\x90\x80\x80"), std::string("\xC3\xB4\xC2\x90\xC2\x80\xC2\x80")); // > U+10FFFF
  CHECK_EQ(toXmlUtf8("\xC3"), std::string("\xC3\x83"));         // truncated sequence

  // Sheet names: nothing is escaped.
  CHECK_EQ(toXmlUtf8("a\rb", false), std::string("a\rb"));
  CHECK_EQ(toXmlUtf8("AB_x0041_7", false), std::string("AB_x0041_7"));
}

void decodeExcelEscapesTest()
{
  CHECK_EQ(decodeExcelEscapes("no escapes"), std::string("no escapes"));
  CHECK_EQ(decodeExcelEscapes("a_x000D_b"), std::string("a\rb"));
  CHECK_EQ(decodeExcelEscapes("_X000a_"), std::string("\n"));
  CHECK_EQ(decodeExcelEscapes("a_x0000_b"), std::string("ab"));
  CHECK_EQ(decodeExcelEscapes("_x005F_"), std::string("_"));
  CHECK_EQ(decodeExcelEscapes("_xFFFF_"), std::string("\xEF\xBF\xBF"));
  // Printable escapes stay verbatim (tools that do not escape store them so).
  CHECK_EQ(decodeExcelEscapes("AB_x0041_7"), std::string("AB_x0041_7"));
  CHECK_EQ(decodeExcelEscapes("_x00G1_"), std::string("_x00G1_"));
  CHECK_EQ(decodeExcelEscapes("_x000D"), std::string("_x000D"));

  const char *texts[] = {
    "a\rb\r\n", "\x01\x02\x1F", "AB_x0041_7", "_x005F_", "__x000D__",
    "\xEF\xBF\xBE", u8"ä€\U0001F600", "tab\tlf\n"
  };
  for (const char *text : texts)
    CHECK_EQ(decodeExcelEscapes(toXmlUtf8(text)), std::string(text));
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

void dateFormatTest()
{
  for (unsigned id : {14u, 15u, 20u, 22u, 27u, 36u, 45u, 47u})
    CHECK(isBuiltinDateFormatId(id));
  for (unsigned id : {0u, 1u, 2u, 13u, 23u, 26u, 37u, 44u, 46u, 49u})
    CHECK(!isBuiltinDateFormatId(id));

  for (const char *code : {"yyyy-mm-dd", "dd.mm.yyyy hh:mm", "hh:mm:ss", "[$-409]mmm yy",
                           "[Red]yyyy", "\"Date:\" d"})
    CHECK(isDateFormatCode(code));
  for (const char *code : {"General", "0.00", "#,##0", "[h]:mm", "[mm]:ss", "[s]", "[Red]0.00",
                           "\"days\" 0", "\\d0", "0.00E+00"})
    CHECK(!isDateFormatCode(code));
}

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
// OpenXLSX behaviour the extension relies on
//------------------------------------------------------------------------------

// Numbers written by OpenXLSX and read back through cellNumber().
void numberRoundTripTest()
{
  TempFile file("numbers");
  {
    XLDocument doc;
    doc.create(file.name(), XLForceOverwrite);
    auto wks = doc.workbook().worksheet(1);
    wks.cell("A1").value() = 87.0;
    wks.cell("A2").value() = 1e20;
    wks.cell("A3").value() = int64_t(5000000000LL);
    wks.cell("A4").value() = int64_t(-7);
    wks.cell("A5").value() = 0.25;
    wks.cell("A6").value() = -3e-7;
    wks.cell("A7").value() = std::string("text");
    wks.cell("A8").value() = toXmlUtf8("cr\rlf\n_x0041_");
    doc.save();
    doc.close();
  }

  XLDocument doc;
  doc.open(file.name());
  auto wks = doc.workbook().worksheet(1);

  CellNumber num;
  CHECK(numberAt(wks, "A1", num));
  CHECK_EQ(num.value, 87.0);

  // "%.17g" writes 1e+20, which OpenXLSX's integer parser reads as 1.
  num = CellNumber();
  CHECK(numberAt(wks, "A2", num));
  CHECK_EQ(num.value, 1e20);
  CHECK(!num.isInteger);

  num = CellNumber();
  CHECK(numberAt(wks, "A3", num));
  CHECK(num.isInteger);
  CHECK_EQ(num.integer, int64_t(5000000000LL));

  num = CellNumber();
  CHECK(numberAt(wks, "A4", num));
  CHECK(num.isInteger);
  CHECK_EQ(num.integer, int64_t(-7));

  num = CellNumber();
  CHECK(numberAt(wks, "A5", num));
  CHECK_EQ(num.value, 0.25);

  num = CellNumber();
  CHECK(numberAt(wks, "A6", num));
  CHECK_EQ(num.value, -3e-7);

  num = CellNumber();
  CHECK(!numberAt(wks, "A7", num));

  XLCell text = wks.cell("A8");
  CHECK_EQ(decodeExcelEscapes(text.value().get<std::string>()), std::string("cr\rlf\n_x0041_"));
  doc.close();
}

// Header texts of cells as written by OpenXLSX, including date formats.
void headerTextTest()
{
  setTimeZone("UTC0");
  TempFile file("headers");
  {
    XLDocument doc;
    doc.create(file.name(), XLForceOverwrite);
    auto wks = doc.workbook().worksheet(1);
    // The format the writer uses; date-only built-in formats (14) are in the fixture.
    XLStyleIndex dateTime = createDateTimeFormat(doc);

    wks.cell("A1").value() = std::string("Name");
    wks.cell("B1").value() = int64_t(2024);
    wks.cell("C1").value() = 2.5;
    wks.cell("D1").value() = true;
    wks.cell("E1").value() = XLDateTime(46023.5);
    wks.cell("E1").setCellFormat(dateTime);
    wks.cell("F1").value() = XLDateTime(46023.0);
    wks.cell("F1").setCellFormat(dateTime); // no time part: date only
    wks.cell("G1").value() = 0.5; // time-only serial stays a number
    wks.cell("G1").setCellFormat(dateTime);
    wks.cell("H1").value() = std::string("a_x000D_b");
    wks.cell("J1").value() = std::string("far");
    doc.save();
    doc.close();
  }

  XLDocument doc;
  doc.open(file.name());
  auto wks = doc.workbook().worksheet(1);
  XLStyles styles = doc.styles();
  std::unordered_map<XLStyleIndex, bool> dateCache;
  ReadState state;

  std::vector<std::string> headers;
  std::vector<bool> exists;
  forEachRow(wks, 1, 1, [&](const SheetRow &row)
  {
    forEachCell(row, wks.columnCount(), [&](uint16_t, XLCell *cell)
    {
      exists.push_back(cell != nullptr);
      headers.push_back(cell ? headerText(*cell, styles, dateCache, state) : std::string());
    });
  });

  CHECK_EQ(headers.size(), size_t(10));
  if (headers.size() == 10)
  {
    CHECK_EQ(headers[0], std::string("Name"));
    CHECK_EQ(headers[1], std::string("2024"));
    CHECK_EQ(headers[2], std::string("2.5"));
    CHECK_EQ(headers[3], std::string("TRUE"));
    CHECK_EQ(headers[4], std::string("2026-01-01 12:00:00"));
    CHECK_EQ(headers[5], std::string("2026-01-01"));
    CHECK_EQ(headers[6], std::string("0.5"));
    CHECK_EQ(headers[7], std::string("a\rb"));
    CHECK(!exists[8]); // I1 was never written
    CHECK_EQ(headers[9], std::string("far"));
  }
  doc.close();
}

void referenceColumnTest()
{
  CHECK_EQ(referenceColumn("A1"), uint16_t(1));
  CHECK_EQ(referenceColumn("Z9"), uint16_t(26));
  CHECK_EQ(referenceColumn("AA1"), uint16_t(27));
  CHECK_EQ(referenceColumn("AZ1"), uint16_t(52));
  CHECK_EQ(referenceColumn("XFD1048576"), uint16_t(16384));
  CHECK_EQ(referenceColumn("XFE1"), uint16_t(0));
  CHECK_EQ(referenceColumn("ZZZZZ1"), uint16_t(0));
  CHECK_EQ(referenceColumn("B"), uint16_t(2));
  CHECK_EQ(referenceColumn("1"), uint16_t(0));
  CHECK_EQ(referenceColumn("a1"), uint16_t(0));
  CHECK_EQ(referenceColumn(""), uint16_t(0));
  CHECK_EQ(referenceColumn(nullptr), uint16_t(0));
}

// Row numbers visited by forEachRow, each followed by its cell pattern
// ('x' = cell, '.' = none), e.g. "1:x.x 3:.x".
std::string rowPattern(XLWorksheet &wks, uint32_t firstRow, uint32_t lastRow, uint16_t colCount)
{
  std::string pattern;
  forEachRow(wks, firstRow, lastRow, [&](const SheetRow &row)
  {
    if (!pattern.empty())
      pattern += ' ';
    pattern += std::to_string(row.number) + (row.hidden ? "h:" : ":");
    forEachCell(row, colCount, [&](uint16_t, XLCell *cell) { pattern += cell ? 'x' : '.'; });
  });
  return pattern;
}

// forEachRow / forEachCell must not create the cells and rows they skip.
void sparseSheetTest()
{
  TempFile file("sparse");
  {
    XLDocument doc;
    doc.create(file.name(), XLForceOverwrite);
    auto wks = doc.workbook().worksheet(1);
    wks.cell("A1").value() = std::string("a");
    wks.cell("C1").value() = std::string("c");
    wks.cell("B500").value() = std::string("far");
    doc.save();
    doc.close();
  }

  XLDocument doc;
  doc.open(file.name());
  auto wks = doc.workbook().worksheet(1);
  CHECK_EQ(wks.rowCount(), uint32_t(500));
  CHECK_EQ(wks.columnCount(), uint16_t(3));

  CHECK_EQ(rowPattern(wks, 1, 1, 3), std::string("1:x.x"));
  CHECK_EQ(rowPattern(wks, 1, 1, 3), std::string("1:x.x")); // B1 still not created
  CHECK_EQ(rowPattern(wks, 2, 499, 3), std::string());
  CHECK_EQ(rowPattern(wks, 500, 500, 3), std::string("500:.x."));
  CHECK_EQ(rowPattern(wks, 1, 500, 1), std::string("1:x 500:.")); // colCount cuts off
  CHECK_EQ(rowPattern(wks, 1, 500, 5), std::string("1:x.x.. 500:.x..."));

  size_t rows = 0;
  for (XMLNode row = sheetDataNode(wks).first_child_of_type(pugi::node_element); !row.empty();
       row = row.next_sibling_of_type(pugi::node_element))
    ++rows;
  CHECK_EQ(rows, size_t(2)); // nothing was created
  CHECK_EQ(wks.rowCount(), uint32_t(500));
  SheetExtent extent = sheetExtent(wks);
  CHECK_EQ(extent.rows, uint32_t(500));
  CHECK_EQ(extent.columns, uint16_t(3));
  doc.close();
}

// The r attributes of rows and cells are optional in the file format: an
// element without one follows the previous one. Elements that do not ascend
// are skipped, other children of a row (<extLst>) ignored.
void implicitReferenceTest()
{
  TempFile file("implicit");
  XLDocument doc;
  doc.create(file.name(), XLForceOverwrite);
  auto wks = doc.workbook().worksheet(1);
  wks.cell("A1").value() = 1;
  wks.cell("B1").value() = 2;
  wks.cell("D1").value() = 4;
  wks.cell("A2").value() = 5;
  wks.cell("C4").value() = 6;
  wks.cell("A6").value() = 7;
  wks.row(6).setHidden(true);

  XMLNode sheetData = sheetDataNode(wks);
  XMLNode row1 = sheetData.first_child_of_type(pugi::node_element);
  XMLNode row2 = row1.next_sibling_of_type(pugi::node_element);
  XMLNode row4 = row2.next_sibling_of_type(pugi::node_element);
  XMLNode cellB1 = row1.first_child_of_type(pugi::node_element).next_sibling_of_type(pugi::node_element);
  XMLNode cellD1 = cellB1.next_sibling_of_type(pugi::node_element);
  cellB1.remove_attribute("r");          // still B1
  cellD1.remove_attribute("r");          // now C1
  row1.append_child("extLst");
  row2.remove_attribute("r");            // still row 2
  row4.attribute("r").set_value(2);      // does not ascend: skipped

  CHECK_EQ(rowPattern(wks, 1, 10, 4), std::string("1:xxx. 2:x... 6h:x..."));
  SheetExtent extent = sheetExtent(wks); // D1 is now C1, row 4 is skipped
  CHECK_EQ(extent.rows, uint32_t(6));
  CHECK_EQ(extent.columns, uint16_t(3));
  CHECK_EQ(rowPattern(wks, 2, 5, 4), std::string("2:x..."));

  std::vector<int64_t> values;
  forEachRow(wks, 1, 1, [&](const SheetRow &row)
  {
    forEachCell(row, 4, [&](uint16_t, XLCell *cell)
    {
      if (cell)
        values.push_back(cell->value().get<int64_t>());
    });
  });
  CHECK(values == std::vector<int64_t>({1, 2, 4}));
  doc.close();
}

void cellReferenceTest()
{
  char buf[16];
  CHECK_EQ(std::string(cellReference(1, 1, buf)), std::string("A1"));
  CHECK_EQ(std::string(cellReference(7, 26, buf)), std::string("Z7"));
  CHECK_EQ(std::string(cellReference(7, 27, buf)), std::string("AA7"));
  CHECK_EQ(std::string(cellReference(12, 702, buf)), std::string("ZZ12"));
  CHECK_EQ(std::string(cellReference(12, 703, buf)), std::string("AAA12"));
  CHECK_EQ(std::string(cellReference(1048576, 16384, buf)), std::string("XFD1048576"));
  for (int c : {1, 2, 25, 26, 27, 52, 53, 701, 702, 703, 16384})
  {
    CHECK_EQ(referenceColumn(cellReference(3, static_cast<uint16_t>(c), buf)), static_cast<uint16_t>(c));
    CHECK_EQ(std::string(cellReference(1048576, static_cast<uint16_t>(c), buf)),
             XLCellReference(1048576, static_cast<uint16_t>(c)).address());
  }
}

// SheetWriter must produce the cells OpenXLSX's own cell.value() = ...
// produces: same type, value, <v> text and format, with shared strings
// reused across sheets.
void sheetWriterTest()
{
  TempFile file("writer");
  const double nan = std::numeric_limits<double>::quiet_NaN();
  const double inf = std::numeric_limits<double>::infinity();
  {
    XLDocument doc;
    doc.create(file.name(), XLForceOverwrite);
    XLStyleIndex dateTime = createDateTimeFormat(doc);
    auto api = doc.workbook().worksheet(1);
    doc.workbook().addWorksheet("Writer");
    auto written = doc.workbook().worksheet("Writer");

    // Row 1: integers, row 2: floats, row 3: booleans and texts.
    const int64_t ints[] = {0, -5, 2147483648LL, std::numeric_limits<int64_t>::max()};
    const double floats[] = {0.1, 123.0, 1e20, -2.5e-300, nan, inf};
    const char *texts[] = {"a", "", " lead", "a", "Ã¤â¬"};

    for (uint16_t c = 0; c < 4; ++c)
      api.cell(1, static_cast<uint16_t>(c + 1)).value() = ints[c];
    for (uint16_t c = 0; c < 6; ++c)
      api.cell(2, static_cast<uint16_t>(c + 1)).value() = floats[c];
    api.cell(3, 1).value() = true;
    api.cell(3, 2).value() = false;
    for (uint16_t c = 0; c < 5; ++c)
      api.cell(3, static_cast<uint16_t>(c + 3)).value() = std::string(texts[c]);
    api.cell(4, 2).value() = XLDateTime(46023.5);
    api.cell(4, 2).setCellFormat(dateTime);

    SharedStringIndex strings(doc.sharedStrings());
    SheetWriter out(written, strings);
    out.beginRow(1);
    for (uint16_t c = 0; c < 4; ++c)
      out.integer(static_cast<uint16_t>(c + 1), ints[c]);
    out.beginRow(2);
    for (uint16_t c = 0; c < 6; ++c)
      out.number(static_cast<uint16_t>(c + 1), floats[c]);
    out.beginRow(3);
    out.boolean(1, true);
    out.boolean(2, false);
    for (uint16_t c = 0; c < 5; ++c)
      out.text(static_cast<uint16_t>(c + 3), texts[c]);
    out.beginRow(4);
    out.number(2, 46023.5, dateTime);

    bool threw = false;
    try { out.beginRow(4); } catch (const std::invalid_argument &) { threw = true; }
    CHECK(threw); // rows must ascend
    threw = false;
    try { out.number(2, 1.0); } catch (const std::invalid_argument &) { threw = true; }
    CHECK(threw); // columns must ascend
    threw = false;
    try { SheetWriter again(written, strings); again.beginRow(4); } catch (const std::invalid_argument &) { threw = true; }
    CHECK(threw); // a new writer starts after the existing rows

    // Five texts, four distinct ones, all already written on the API sheet.
    CHECK_EQ(doc.sharedStrings().stringCount(), int32_t(4));

    // Writers sharing one index see each other's new strings.
    doc.workbook().addWorksheet("Second");
    doc.workbook().addWorksheet("Third");
    auto second = doc.workbook().worksheet("Second");
    auto third = doc.workbook().worksheet("Third");
    SheetWriter w2(second, strings);
    SheetWriter w3(third, strings);
    w2.beginRow(1);
    w3.beginRow(1);
    w2.text(1, "new");
    w3.text(1, "new");
    CHECK_EQ(doc.sharedStrings().stringCount(), int32_t(5));
    CHECK_EQ(strings.indexOf("new"), doc.sharedStrings().getStringIndex("new"));
    CHECK_EQ(strings.indexOf("a"), doc.sharedStrings().getStringIndex("a"));
    doc.save();
    doc.close();
  }

  XLDocument doc;
  doc.open(file.name());
  auto api = doc.workbook().worksheet(1);
  auto written = doc.workbook().worksheet("Writer");
  for (const char *ref : {"A1", "B1", "C1", "D1", "A2", "B2", "C2", "D2", "E2", "F2",
                          "A3", "B3", "C3", "D3", "E3", "F3", "G3", "A4", "B4"})
  {
    XLCell expected = api.findCell(ref);
    XLCell actual = written.findCell(ref);
    CHECK_EQ(bool(actual), bool(expected));
    if (!expected || !actual)
      continue;
    CHECK(actual.value().type() == expected.value().type());
    CHECK_EQ(actual.value().getString(), expected.value().getString());
    CHECK_EQ(rawNumberText(actual), rawNumberText(expected));
    CHECK_EQ(actual.cellFormat(), expected.cellFormat());
  }
  CHECK_EQ(written.findCell("G3").value().get<std::string>(), std::string("Ã¤â¬"));
  CHECK_EQ(written.findCell("E2").value().getString(), std::string("#NUM!"));
  CHECK_EQ(written.findCell("D1").value().get<int64_t>(), std::numeric_limits<int64_t>::max());
  doc.close();
}

// Cells from testdata/CtrlExcelReaderFixture.xlsx (written like Excel does,
// see testdata/make_fixture.py), column B.
void fixtureTest()
{
  if (!std::filesystem::exists(std::filesystem::u8path(fixturePath)))
  {
    report(false, "fixture not found: '" + fixturePath + "'", __FILE__, __LINE__);
    return;
  }

  setTimeZone("UTC0");
  XLDocument doc;
  doc.open(fixturePath);
  auto wks = doc.workbook().worksheet("Fixture");
  XLStyles styles = doc.styles();
  std::unordered_map<XLStyleIndex, bool> dateCache;
  ReadState state;
  const time_t maxSec = std::numeric_limits<time_t>::max();

  // pre1970: 1960-05-01 is a date, outside CTRL time.
  {
    XLCell cell = wks.cell("B2");
    CellNumber num;
    CHECK(numberAt(wks, "B2", num));
    CHECK(isDateNumber(num, cell, styles, dateCache));
    std::tm tm{};
    int16_t milli = 0;
    time_t sec = 0;
    CHECK(serialToLocalTime(num.value, tm, milli, sec, maxSec, state) == LocalTime::OutOfRange);
    CHECK_EQ(headerText(cell, styles, dateCache, state), std::string("1960-05-01"));
  }
  // dateInt: whole-number serial, typed Integer, still a date.
  {
    XLCell cell = wks.cell("B3");
    CellNumber num;
    CHECK(numberAt(wks, "B3", num));
    CHECK(num.isInteger);
    CHECK(isDateNumber(num, cell, styles, dateCache));
    CHECK_EQ(headerText(cell, styles, dateCache, state), std::string("2026-01-01"));
  }
  // duration: [h]:mm is a number.
  {
    XLCell cell = wks.cell("B4");
    CellNumber num;
    CHECK(numberAt(wks, "B4", num));
    CHECK(!isDateNumber(num, cell, styles, dateCache));
    CHECK_EQ(num.value, 1.5);
  }
  // error: the code is kept in <v>.
  {
    XLCell cell = wks.cell("B5");
    CHECK(cell.value().type() == XLValueType::Error);
    CHECK_EQ(rawNumberText(cell), std::string("#DIV/0!"));
    CHECK(!isUncachedFormula(cell, cell.value()));
  }
  // uncached: formula without <v>.
  {
    XLCell cell = wks.cell("B6");
    CHECK(isUncachedFormula(cell, cell.value()));
    CHECK_EQ(headerText(cell, styles, dateCache, state), std::string());
  }
  // crlf / literal: only Excel's own escapes are decoded.
  CHECK_EQ(decodeExcelEscapes(wks.cell("B7").value().get<std::string>()), std::string("a\r\nb"));
  CHECK_EQ(decodeExcelEscapes(wks.cell("B8").value().get<std::string>()), std::string("AB_x0041_7"));
  // exponent: "1E+20" without '.', which OpenXLSX reads as Integer 1.
  {
    XLCell cell = wks.cell("B9");
    CellNumber num;
    CHECK(numberAt(wks, "B9", num));
    CHECK(!num.isInteger);
    CHECK_EQ(num.value, 1e20);
    CHECK_EQ(headerText(cell, styles, dateCache, state), std::string("1e+20"));
  }
  // bigint: beyond int64, which OpenXLSX clamps.
  {
    CellNumber num;
    CHECK(numberAt(wks, "B10", num));
    CHECK(!num.isInteger);
    CHECK_EQ(num.value, 12345678901234567890.0);
  }
  // Row 11 only carries formatting; it exists but has no cells.
  {
    CHECK_EQ(rowPattern(wks, 11, 11, 2), std::string("11:.."));
  }
  // Sheet "Implicit": row 3 missing, rows 5 and 6 and some cells without r,
  // row 5 ends in <extLst> (which made OpenXLSX's noexcept columnCount() throw).
  {
    auto implicit = doc.workbook().worksheet("Implicit");
    SheetExtent extent = sheetExtent(implicit);
    CHECK_EQ(extent.rows, uint32_t(6));
    CHECK_EQ(extent.columns, uint16_t(2));
    CHECK_EQ(rowPattern(implicit, 1, 6, 2), std::string("1:xx 2:xx 4:xx 5:xx 6:xx"));

    std::vector<int64_t> values;
    forEachRow(implicit, 2, 6, [&](const SheetRow &row)
    {
      forEachCell(row, 2, [&](uint16_t c, XLCell *cell)
      {
        if (c == 1 && cell)
          values.push_back(cell->value().get<int64_t>());
      });
    });
    CHECK(values == std::vector<int64_t>({1, 2, 3, 4}));
  }
  doc.close();
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
    {"toXmlUtf8Test", toXmlUtf8Test},
    {"decodeExcelEscapesTest", decodeExcelEscapesTest},
    {"foldCaseTest", foldCaseTest},
    {"checkSheetNamesTest", checkSheetNamesTest},
    {"numberWidthTest", numberWidthTest},
    {"dateFormatTest", dateFormatTest},
    {"utcTimeTest", utcTimeTest},
#ifndef _WIN32
    {"dstTimeTest", dstTimeTest},
#endif
    {"numberRoundTripTest", numberRoundTripTest},
    {"headerTextTest", headerTextTest},
    {"referenceColumnTest", referenceColumnTest},
    {"sparseSheetTest", sparseSheetTest},
    {"implicitReferenceTest", implicitReferenceTest},
    {"cellReferenceTest", cellReferenceTest},
    {"sheetWriterTest", sheetWriterTest},
    {"fixtureTest", fixtureTest},
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
