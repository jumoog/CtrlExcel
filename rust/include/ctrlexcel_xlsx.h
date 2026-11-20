/* C ABI of the Rust library rust/ (calamine and rust_xlsxwriter) that reads
 * and writes .xlsx files for CtrlExcelReader. Implemented in rust/src/ffi.rs;
 * keep both in sync. ExcelXlsxCore.hxx wraps it for C++.
 *
 * Texts are UTF-8 with an explicit length (not NUL-terminated). Rows and
 * columns are 1-based. Functions that can fail take a cxlsx_error, whose set()
 * receives the reason (UTF-8) before the function returns false (or NULL, or
 * CXLSX_FAILED). No function throws or unwinds: Rust panics are caught and
 * reported as errors. Callbacks must not throw either. */

#ifndef CTRLEXCEL_XLSX_H
#define CTRLEXCEL_XLSX_H

#include <stdbool.h>
#include <stddef.h>
#include <stdint.h>

#ifdef __cplusplus
extern "C" {
#endif

typedef struct cxlsx_error
{
  void *ctx;
  void (*set)(void *ctx, const char *message, size_t len);
} cxlsx_error;

/* ---- Reading --------------------------------------------------------------- */

typedef struct cxlsx_book cxlsx_book;

enum cxlsx_cell_type
{
  CXLSX_BOOL = 1,
  CXLSX_NUMBER = 2,
  CXLSX_DATETIME = 3,        /* number: serial in the 1900 date system */
  CXLSX_DURATION = 4,        /* number: days ([h]:mm and similar formats) */
  CXLSX_TEXT = 5,            /* text: escapes (_xHHHH_) decoded, no NUL */
  CXLSX_ERROR = 6,           /* text: Excel's error code, e.g. "#DIV/0!" */
  CXLSX_UNCACHED_FORMULA = 7 /* a formula never calculated: no value */
};

typedef struct cxlsx_cell
{
  uint32_t row;
  uint32_t column;
  uint32_t type; /* cxlsx_cell_type */
  bool boolean;
  double number;
  const char *text; /* valid during the callback only */
  size_t text_len;
} cxlsx_cell;

/* begin() once; then row() for every row stored in the sheet, ascending, each
 * followed by its cells (ascending columns; cells without a value are left
 * out). Returning false stops the read (CXLSX_STOPPED). */
typedef struct cxlsx_sheet_visitor
{
  void *ctx;
  bool (*begin)(void *ctx, uint32_t rows, uint32_t columns); /* last row, highest column */
  bool (*row)(void *ctx, uint32_t row, bool hidden);
  bool (*cell)(void *ctx, const cxlsx_cell *cell);
} cxlsx_sheet_visitor;

enum cxlsx_status
{
  CXLSX_OK = 0,
  CXLSX_FAILED = 1,
  CXLSX_STOPPED = 2
};

/* Reads the whole file into memory. NULL on failure. */
cxlsx_book *cxlsx_book_open(const char *path, size_t path_len, const cxlsx_error *error);
void cxlsx_book_close(cxlsx_book *book);

/* Worksheets (not chart or dialog sheets) in workbook order. The name stays
 * valid until the book is closed. */
size_t cxlsx_book_sheet_count(const cxlsx_book *book);
const char *cxlsx_book_sheet_name(const cxlsx_book *book, size_t index, size_t *len);

/* Returns a cxlsx_status. */
int cxlsx_book_read_sheet(cxlsx_book *book, const char *name, size_t name_len,
                          const cxlsx_sheet_visitor *visitor, const cxlsx_error *error);

/* ---- Writing --------------------------------------------------------------- */

typedef struct cxlsx_writer cxlsx_writer;

enum cxlsx_style
{
  CXLSX_STYLE_DEFAULT = 0,
  CXLSX_STYLE_HEADER = 1,  /* bold */
  CXLSX_STYLE_DATETIME = 2 /* built-in number format 22, "m/d/yyyy h:mm" */
};

/* NULL only if out of memory. */
cxlsx_writer *cxlsx_writer_new(void);
void cxlsx_writer_free(cxlsx_writer *writer);

/* The cell functions write to the sheet added last. Within a sheet, rows must
 * be written in ascending order (finished rows go to a temporary file). */
bool cxlsx_writer_add_sheet(cxlsx_writer *writer, const char *name, size_t name_len,
                            const cxlsx_error *error);
bool cxlsx_writer_text(cxlsx_writer *writer, uint32_t row, uint32_t column,
                       const char *text, size_t text_len, uint32_t style, const cxlsx_error *error);
/* Non-finite values are written as the error #NUM!. */
bool cxlsx_writer_number(cxlsx_writer *writer, uint32_t row, uint32_t column,
                         double value, uint32_t style, const cxlsx_error *error);
bool cxlsx_writer_bool(cxlsx_writer *writer, uint32_t row, uint32_t column,
                       bool value, uint32_t style, const cxlsx_error *error);
/* Width in characters, capped at 255. */
bool cxlsx_writer_column_width(cxlsx_writer *writer, uint32_t column, double width,
                               const cxlsx_error *error);

/* Writes to a temporary file next to path and renames it over path. A
 * workbook without sheets gets an empty "Sheet1". Call at most once. */
bool cxlsx_writer_save(cxlsx_writer *writer, const char *path, size_t path_len,
                       const cxlsx_error *error);

#ifdef __cplusplus
}
#endif

#endif
