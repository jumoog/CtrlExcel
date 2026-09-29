#ifndef _EXCEL_XLSX_CORE_HXX_
#define _EXCEL_XLSX_CORE_HXX_

// Helpers that need only OpenXLSX and the standard library, not the WinCC OA
// API, so tests/ExcelXlsxCoreTest.cxx can build and run them without it
// (e.g. in CI). Texts are UTF-8 throughout.

#include <climits>
#include <cstddef>
#include <cstdint>
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

  // Call fn(columnIndex, cell) for columns 1..colCount of an existing row,
  // with cell == nullptr where the row has no cell. Dereferencing OpenXLSX
  // cell iterators creates missing cells in the XML (and rows().cells()
  // visits every position), which for a sheet with a stray cell far out
  // (row 1048576, column XFD) would materialise millions of nodes.
  template <typename Fn>
  void forEachCell(OpenXLSX::XLWorksheet &wks, uint32_t row, uint16_t colCount, Fn &&fn)
  {
    OpenXLSX::XLCellRange range = wks.range(OpenXLSX::XLCellReference(row, 1),
                                            OpenXLSX::XLCellReference(row, colCount));
    uint16_t c = 0;
    for (auto it = range.begin(); it != range.end(); ++it, ++c)
      fn(c, it.cellExists() ? &*it : nullptr);
  }

  // names are UTF-8; converts them to valid UTF-8 in place. Returns an empty
  // string if Excel accepts every name, otherwise the first problem found.
  std::string checkSheetNames(std::vector<std::string> &names);
}

#endif
