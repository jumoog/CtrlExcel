#ifndef _EXCEL_XLSX_CORE_HXX_
#define _EXCEL_XLSX_CORE_HXX_

// Helpers that need only OpenXLSX and the standard library, not the WinCC OA
// API, so tests/ExcelXlsxCoreTest.cxx can build and run them without it
// (e.g. in CI). Texts are UTF-8 throughout.

#include <climits>
#include <cstddef>
#include <cstdint>
#include <cstring>
#include <ctime>
#include <string>
#include <unordered_map>
#include <vector>

// WinCC OA's win32.h typedefs ssize_t as int; OpenXLSX redefines it as a
// 64-bit alias. Rename OpenXLSX's alias (no effect on mangled names). Windows
// only: elsewhere there is no clash, and the macro could rename a system
// header's ssize_t first included through OpenXLSX.
#ifdef _WIN32
#define ssize_t OpenXLSX_ssize_t
#endif
#include <OpenXLSX.hpp>
#include <headers/XLXmlParser.hpp> // XMLNode, for forEachRow / forEachCell
#ifdef _WIN32
#undef ssize_t
#endif

namespace ExcelXlsxCore
{
  constexpr long long DAY_MILLIS = 24LL * 60LL * 60LL * 1000LL;
  constexpr long long DAY_SECONDS = 24LL * 60LL * 60LL;
  constexpr size_t EXCEL_MAX_CELL_CHARS = 32767;
  constexpr uint32_t EXCEL_FMT_DATE_TIME = 22; // built-in "m/d/yyyy h:mm"

  // Excel serial day of 1970-01-01. Serials below 61 precede Excel's
  // fictitious 1900-02-29 and are shifted by one day.
  constexpr long long EXCEL_UNIX_EPOCH_DAY = 25569;

  // Days since 1970-01-01 of a proleptic Gregorian date (H. Hinnant's
  // days_from_civil). Pure arithmetic, independent of the C runtime.
  long long daysFromCivil(long long y, unsigned m, unsigned d);

  // Inverse of daysFromCivil.
  void civilFromDays(long long z, long long &y, unsigned &m, unsigned &d);

  long long excelDayToUnixDay(long long excelDay);
  long long unixDayToExcelDay(long long unixDay);

  // UTF-16 length of a UTF-8 string, which is what Excel's limits count.
  size_t utf16Length(const std::string &s);

  // Text (UTF-8) as valid XML 1.0 UTF-8 for a cell, which Excel requires.
  // Characters XML cannot hold (control characters other than tab and LF,
  // CR, which XML parsers normalise to LF, and U+FFFE/U+FFFF) are written as
  // Excel's "_xHHHH_" escapes, and a literal "_xHHHH_" in the text gets its
  // '_' escaped, so everything round-trips through decodeExcelEscapes().
  // Bytes that are not valid UTF-8 are taken as ISO-8859-1 as a last resort.
  // For sheet names (cellText == false) nothing is escaped; checkSheetNames
  // rejects control characters there instead.
  std::string toXmlUtf8(const char *text, bool cellText = true);

  // Decode Excel's "_xHHHH_" escapes in cell text, which OpenXLSX leaves as
  // literal text. Only the code units Excel itself escapes are decoded:
  // control characters, '_' (the escape of a literal "_xHHHH_") and
  // U+FFFE/U+FFFF. Other sequences are kept verbatim, since tools that do
  // not escape store such text literally (e.g. a part number "AB_x0041_7").
  // U+0000 is dropped: it would end the C string and cut off the rest.
  std::string decodeExcelEscapes(const std::string &text);

  // Code points of a valid UTF-8 string with a simple case folding (ASCII,
  // Latin-1, Latin Extended-A, Greek, Cyrillic), enough to catch the sheet
  // names Excel treats as duplicates.
  std::u32string foldCase(const std::string &utf8);

  // Displayed width of a number, for column sizing.
  size_t numberWidth(double value);

  // New cell formats, copies of the default one: bold (header row) and date +
  // time display (EXCEL_FMT_DATE_TIME). Return the new style index.
  OpenXLSX::XLStyleIndex createBoldHeaderFormat(OpenXLSX::XLDocument &doc);
  OpenXLSX::XLStyleIndex createDateTimeFormat(OpenXLSX::XLDocument &doc);

  // Built-in Excel number-format IDs that represent dates/times (ECMA-376).
  // 46 ("[h]:mm:ss") is an elapsed duration, not a point in time.
  bool isBuiltinDateFormatId(unsigned int id);

  // Scan a custom format-code string for date/time tokens (y m d h s)
  // while ignoring quoted literals, escaped chars and bracketed sections.
  // Elapsed-time formats ([h]:mm, [mm]:ss, [s]) are durations, not dates.
  bool isDateFormatCode(const std::string &code);

  // Whether the cell's number format is a date/time format. Results are
  // cached per cell-format index so each distinct format is inspected once.
  bool isDateCell(OpenXLSX::XLCell &cell, const OpenXLSX::XLStyles &styles,
                  std::unordered_map<OpenXLSX::XLStyleIndex, bool> &dateCache);

  // State of one sheet read: counters for cells read with a fallback
  // (reported as warnings afterwards) and a cache of each calendar day's UTC
  // offset. mktime is slow on MSVC (it re-evaluates the time-zone rules on
  // every call), so it runs about twice per distinct day instead of once per
  // date cell; only days on which the offset changes (DST transitions) are
  // converted per cell.
  struct ReadState
  {
    static constexpr long long OFFSET_VARIES = LLONG_MIN;

    size_t uncachedFormulas = 0;     // read as ""
    size_t unrepresentableDates = 0; // read as the Excel serial (float)
    size_t shiftedTimes = 0;         // nonexistent local times, moved by mktime
    std::unordered_map<long long, long long> dayOffset; // unixDay -> seconds east of UTC
  };

  enum class LocalTime
  {
    Ok,
    Shifted,    // did not exist (hour skipped at the DST spring-forward)
    OutOfRange  // before 1970-01-01 UTC or after maxSec
  };

  // Convert an Excel date serial (local wall-clock time) to epoch seconds and
  // milliseconds; tm receives the calendar fields. CTRL time holds only
  // 1970-01-01 UTC up to TimeVar::MaxTimeVarSec (2262 on 64-bit), which the
  // caller passes as maxSec, so e.g. a 1960 birth date or the 1900-01-01
  // placeholder is OutOfRange.
  LocalTime serialToLocalTime(double serial, std::tm &tm, int16_t &milli, time_t &sec,
                              time_t maxSec, ReadState &state);

  // Epoch seconds and milliseconds to the Excel serial of the local
  // wall-clock time (Excel serials have no timezone). Inverse of
  // serialToLocalTime. Every CTRL time (1970..2262) lies within Excel's date
  // range, so only the C runtime's conversion can fail (e.g. MSVC for the
  // first hours of 1970 in zones west of UTC).
  bool localTimeToSerial(time_t sec, int milli, double &serial);

  // Numeric value of a cell. OpenXLSX types numeric text without '.' as
  // Integer, even when it is a date serial or uses an exponent.
  struct CellNumber
  {
    bool isInteger = false; // no fractional part in the stored text
    int64_t integer = 0;    // valid if isInteger
    double value = 0.0;
  };

  // Raw text of the cell's <v> element (OpenXLSX offers no other access).
  std::string rawNumberText(const OpenXLSX::XLCell &cell);

  // False for non-numeric cells.
  bool cellNumber(const OpenXLSX::XLCell &cell, const OpenXLSX::XLCellValue &val, CellNumber &out);

  // Whether a numeric cell holds a date. XLDateTime rejects serials below 1.0
  // (time-only values such as 08:00), so those stay plain numbers.
  bool isDateNumber(const CellNumber &num, OpenXLSX::XLCell &cell, const OpenXLSX::XLStyles &styles,
                    std::unordered_map<OpenXLSX::XLStyleIndex, bool> &dateCache);

  // A formula whose result was never calculated (files generated by tools
  // and not saved by Excel). OpenXLSX types such a cell as Integer and reads
  // the missing or empty <v> as 0, so check the raw value text. (OpenXLSX's
  // own placeholder <v>0</v> cannot be told apart from a real 0.)
  bool isUncachedFormula(const OpenXLSX::XLCell &cell, const OpenXLSX::XLCellValue &val);

  // Header cells may hold numbers, dates or booleans; get<std::string> throws
  // for those, and getString() renders 2024 as "2024.000000".
  std::string headerText(OpenXLSX::XLCell &cell, const OpenXLSX::XLStyles &styles,
                         std::unordered_map<OpenXLSX::XLStyleIndex, bool> &dateCache,
                         ReadState &state);

  // Column number (1-based) of a cell reference such as "AB12"; 0 if it does
  // not start with a column (uppercase letters) of at most XFD.
  uint16_t referenceColumn(const char *ref);

  // The worksheet's <sheetData> element and the shared strings its cells refer
  // to. OpenXLSX keeps the sheet XML protected; see forEachRow for why the
  // reader walks it directly.
  OpenXLSX::XMLNode sheetDataNode(OpenXLSX::XLWorksheet &wks);
  const OpenXLSX::XLSharedStrings &sheetSharedStrings(OpenXLSX::XLWorksheet &wks);

  // A <row> element of the sheet XML.
  struct SheetRow
  {
    OpenXLSX::XMLNode node;
    const OpenXLSX::XLSharedStrings *sharedStrings;
    uint32_t number; // 1-based
    bool hidden;
  };

  // Call fn(const SheetRow &) for every row stored in the sheet with
  // firstRow <= number <= lastRow, in order; missing rows are neither visited
  // nor created. OpenXLSX locates a row by walking the row list from its
  // start or end on every lookup (wks.range(), wks.cell(), the first step of
  // every iterator), so reading a sheet row by row that way is quadratic
  // (100000 rows took 75 s). Its iterators also create the rows and cells
  // they visit, which for a stray cell far out (row 1048576, column XFD)
  // would materialise millions of nodes. One pass over the XML avoids both.
  // Rows are stored in ascending order (OpenXLSX relies on that as well); a
  // row without the optional r attribute follows the previous one, and a row
  // that does not ascend is skipped.
  template <typename Fn>
  void forEachRow(OpenXLSX::XLWorksheet &wks, uint32_t firstRow, uint32_t lastRow, Fn &&fn)
  {
    const OpenXLSX::XLSharedStrings &strings = sheetSharedStrings(wks);
    uint32_t number = 0;
    for (OpenXLSX::XMLNode node = sheetDataNode(wks).first_child_of_type(pugi::node_element);
         !node.empty(); node = node.next_sibling_of_type(pugi::node_element))
    {
      uint32_t r = node.attribute("r").as_uint(0);
      if (r == 0)
        r = number + 1;
      else if (r <= number)
        continue;
      number = r;
      if (number < firstRow)
        continue;
      if (number > lastRow)
        break;
      fn(SheetRow{node, &strings, number, node.attribute("hidden").as_bool(false)});
    }
  }

  // Call fn(column, cellNode) for the <c> elements of a row, in order, with
  // their 1-based column, until fn returns false. A cell without the optional
  // r attribute follows the previous one; one that does not ascend is
  // skipped, as is everything after column XFD.
  template <typename Fn>
  void forEachCellNode(const SheetRow &row, Fn &&fn)
  {
    unsigned previous = 0; // column of the previous <c>
    for (OpenXLSX::XMLNode node = row.node.first_child_of_type(pugi::node_element);
         !node.empty(); node = node.next_sibling_of_type(pugi::node_element))
    {
      if (std::strcmp(node.name(), "c") != 0) // e.g. <extLst>
        continue;
      unsigned column = referenceColumn(node.attribute("r").value());
      if (column == 0)
        column = previous + 1;
      else if (column <= previous)
        continue;
      if (column > OpenXLSX::MAX_COLS)
        break;
      previous = column;
      if (!fn(static_cast<uint16_t>(column), node))
        break;
    }
  }

  // Call fn(columnIndex, cell) for columns 1..colCount of a row, with
  // cell == nullptr where the row has no cell (columnIndex is 0-based).
  // Linear in the row's cells, creates nothing.
  template <typename Fn>
  void forEachCell(const SheetRow &row, uint16_t colCount, Fn &&fn)
  {
    uint16_t done = 0; // columns handed to fn
    forEachCellNode(row, [&](uint16_t column, const OpenXLSX::XMLNode &node)
    {
      if (column > colCount)
        return false;
      while (++done < column)
        fn(static_cast<uint16_t>(done - 1), static_cast<OpenXLSX::XLCell *>(nullptr));
      OpenXLSX::XLCell cell(node, *row.sharedStrings);
      fn(static_cast<uint16_t>(done - 1), &cell);
      return done < colCount;
    });
    while (done < colCount)
    {
      ++done;
      fn(static_cast<uint16_t>(done - 1), static_cast<OpenXLSX::XLCell *>(nullptr));
    }
  }

  // Number of the last row and highest column of a sheet, by the rules of
  // forEachRow / forEachCellNode. Replaces XLWorksheet::rowCount() (the r of
  // the last row: 0 if that row omits it) and columnCount(), which is
  // noexcept yet throws, i.e. terminates the process, for a row whose last
  // element is a <c> without r or an <extLst>.
  struct SheetExtent
  {
    uint32_t rows = 0;
    uint16_t columns = 0;
  };
  SheetExtent sheetExtent(OpenXLSX::XLWorksheet &wks);

  // names are UTF-8; converts them to valid UTF-8 in place. Returns an empty
  // string if Excel accepts every name, otherwise the first problem found.
  std::string checkSheetNames(std::vector<std::string> &names);
}

#endif
