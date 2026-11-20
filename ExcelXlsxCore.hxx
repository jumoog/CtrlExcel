#ifndef _EXCEL_XLSX_CORE_HXX_
#define _EXCEL_XLSX_CORE_HXX_

// Everything that needs the xlsx library (rust/, through its C ABI) and the
// standard library, but not the WinCC OA API, so tests/ExcelXlsxCoreTest.cxx
// can build and run it without the API (e.g. in CI). Texts are UTF-8
// throughout.

#include <ctrlexcel_xlsx.h>

#include <climits>
#include <cstddef>
#include <cstdint>
#include <ctime>
#include <string>
#include <unordered_map>
#include <vector>

namespace ExcelXlsxCore
{
  constexpr long long DAY_MILLIS = 24LL * 60LL * 60LL * 1000LL;
  constexpr long long DAY_SECONDS = 24LL * 60LL * 60LL;
  constexpr size_t EXCEL_MAX_CELL_CHARS = 32767;
  constexpr uint32_t EXCEL_MAX_ROWS = 1048576;
  constexpr uint32_t EXCEL_MAX_COLUMNS = 16384;
  constexpr double EXCEL_MAX_COLUMN_WIDTH = 255.0;

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

  // text as valid UTF-8: bytes that are not valid UTF-8 are taken as
  // ISO-8859-1 as a last resort. The xlsx library escapes what XML cannot
  // hold (control characters, CR, U+FFFE/U+FFFF, a literal "_xHHHH_") as
  // Excel's "_xHHHH_" escapes, and decodes all of them on read.
  std::string toValidUtf8(const char *text);

  // Code points of a valid UTF-8 string with a simple case folding (ASCII,
  // Latin-1, Latin Extended-A, Greek, Cyrillic), enough to catch the sheet
  // names Excel treats as duplicates.
  std::u32string foldCase(const std::string &utf8);

  // Displayed width of a number, for column sizing.
  size_t numberWidth(double value);

  // How a number read from a file fits a CTRL integer type.
  enum class IntegerFit
  {
    None, // fractional, non-finite or beyond int64: float
    Int,  // whole, 32-bit
    Long  // whole, 64-bit
  };

  // Excel stores every number as a double: whole numbers become int (or
  // long beyond 32 bits), the rest float. -0.0 counts as 0.
  IntegerFit integerFit(double value, int64_t &integer);

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

  // Text of a CXLSX_TEXT or CXLSX_ERROR cell.
  inline std::string cellText(const cxlsx_cell &cell)
  {
    return std::string(cell.text ? cell.text : "", cell.text_len);
  }

  // Whether a cell holds a point in time. Serials below 1.0 (time-only
  // values such as 08:00, which would be 1900-01-00) stay plain numbers.
  bool isDateCell(const cxlsx_cell &cell);

  // Text of a header cell, which may also hold a number, date or boolean:
  // whole numbers without decimals ("2024"), dates as "YYYY-MM-DD" (plus
  // " HH:MM:SS" with a time part). Empty for errors and uncalculated formulas.
  std::string headerText(const cxlsx_cell &cell, ReadState &state);

  // names are UTF-8; converts them to valid UTF-8 in place. Returns an empty
  // string if Excel accepts every name, otherwise the first problem found.
  std::string checkSheetNames(std::vector<std::string> &names);

  //----------------------------------------------------------------------------
  // The xlsx library (rust/include/ctrlexcel_xlsx.h). Its functions report
  // failures as std::runtime_error with the library's message (UTF-8).
  //----------------------------------------------------------------------------

  // Receives a sheet from Workbook::readSheet: begin() once, then row() for
  // every row stored in the sheet in ascending order (also rows without
  // cells), each followed by its cells with a value in ascending columns.
  // Returning false ends the read early (not an error). An exception thrown
  // here ends the read and is rethrown by readSheet.
  class SheetVisitor
  {
  public:
    virtual ~SheetVisitor() = default;
    // rows: number of the last row; columns: highest column (0 if empty)
    virtual bool begin(uint32_t rows, uint32_t columns) = 0;
    virtual bool row(uint32_t row, bool hidden) = 0;
    virtual bool cell(const cxlsx_cell &cell) = 0;
  };

  // An .xlsx file opened for reading (read into memory as a whole). One
  // thread at a time.
  class Workbook
  {
  public:
    explicit Workbook(const std::string &path);
    ~Workbook();
    Workbook(const Workbook &) = delete;
    Workbook &operator=(const Workbook &) = delete;

    // Worksheets (not chart or dialog sheets) in workbook order.
    std::vector<std::string> sheetNames() const;

    void readSheet(const std::string &name, SheetVisitor &visitor);

  private:
    cxlsx_book *m_book;
  };

  // Cell formats of WorkbookWriter.
  enum class CellStyle : uint32_t
  {
    Default = CXLSX_STYLE_DEFAULT,
    Header = CXLSX_STYLE_HEADER,    // bold
    DateTime = CXLSX_STYLE_DATETIME // built-in format 22, "m/d/yyyy h:mm"
  };

  // A new workbook, written to a file by save() only: it writes a temporary
  // file next to the target and renames it over the target, so a failed
  // write leaves an existing file untouched. Cells go to the sheet added
  // last, rows in ascending order.
  class WorkbookWriter
  {
  public:
    WorkbookWriter();
    ~WorkbookWriter();
    WorkbookWriter(const WorkbookWriter &) = delete;
    WorkbookWriter &operator=(const WorkbookWriter &) = delete;

    void addSheet(const std::string &name);
    void text(uint32_t row, uint32_t column, const std::string &utf8, CellStyle style = CellStyle::Default);
    // Non-finite values are written as the error #NUM!.
    void number(uint32_t row, uint32_t column, double value, CellStyle style = CellStyle::Default);
    void boolean(uint32_t row, uint32_t column, bool value, CellStyle style = CellStyle::Default);
    // Width in characters, capped at 255.
    void columnWidth(uint32_t column, double width);
    // A workbook without sheets gets an empty "Sheet1".
    void save(const std::string &path);

  private:
    cxlsx_writer *m_writer;
  };
}

#endif
