//! The C ABI declared in include/ctrlexcel_xlsx.h. Every function catches
//! panics: unwinding into the CTRL manager would terminate it.
//!
//! The functions taking pointers are `unsafe`: the header states what the
//! caller guarantees (valid handles, buffers of the given length).

#![allow(clippy::missing_safety_doc)] // see include/ctrlexcel_xlsx.h

use crate::reader::{Book, Cell, CellValue, ReadError, SheetVisitor};
use crate::writer::{Style, Writer};
use std::ffi::{c_char, c_int, c_void};
use std::panic::{catch_unwind, AssertUnwindSafe};
use std::path::Path;

#[repr(C)]
pub struct cxlsx_error {
    ctx: *mut c_void,
    set: Option<extern "C" fn(*mut c_void, *const c_char, usize)>,
}

#[repr(C)]
pub struct cxlsx_cell {
    row: u32,
    column: u32,
    kind: u32,
    boolean: bool,
    number: f64,
    text: *const c_char,
    text_len: usize,
}

#[repr(C)]
pub struct cxlsx_sheet_visitor {
    ctx: *mut c_void,
    begin: Option<extern "C" fn(*mut c_void, u32, u32) -> bool>,
    row: Option<extern "C" fn(*mut c_void, u32, bool) -> bool>,
    cell: Option<extern "C" fn(*mut c_void, *const cxlsx_cell) -> bool>,
}

const CXLSX_BOOL: u32 = 1;
const CXLSX_NUMBER: u32 = 2;
const CXLSX_DATETIME: u32 = 3;
const CXLSX_DURATION: u32 = 4;
const CXLSX_TEXT: u32 = 5;
const CXLSX_ERROR: u32 = 6;
const CXLSX_UNCACHED_FORMULA: u32 = 7;

const CXLSX_OK: c_int = 0;
const CXLSX_FAILED: c_int = 1;
const CXLSX_STOPPED: c_int = 2;

fn report(error: *const cxlsx_error, message: &str) {
    // SAFETY: the caller passes NULL or a valid cxlsx_error.
    if let Some(error) = unsafe { error.as_ref() } {
        if let Some(set) = error.set {
            set(error.ctx, message.as_ptr().cast(), message.len());
        }
    }
}

fn panic_message(payload: &(dyn std::any::Any + Send)) -> String {
    let detail = payload
        .downcast_ref::<&str>()
        .map(|s| s.to_string())
        .or_else(|| payload.downcast_ref::<String>().cloned())
        .unwrap_or_default();
    format!("internal error in the xlsx library: {detail}")
}

/// Runs f, reporting its error or panic; fallback is returned then.
fn guarded<T>(error: *const cxlsx_error, fallback: T, f: impl FnOnce() -> Result<T, String>) -> T {
    match catch_unwind(AssertUnwindSafe(f)) {
        Ok(Ok(value)) => value,
        Ok(Err(message)) => {
            report(error, &message);
            fallback
        }
        Err(payload) => {
            report(error, &panic_message(payload.as_ref()));
            fallback
        }
    }
}

/// SAFETY: ptr must point to len readable bytes (or len be 0).
unsafe fn text_arg<'a>(ptr: *const c_char, len: usize, what: &str) -> Result<&'a str, String> {
    if len == 0 {
        return Ok("");
    }
    if ptr.is_null() {
        return Err(format!("{what} is NULL"));
    }
    let bytes = std::slice::from_raw_parts(ptr.cast::<u8>(), len);
    std::str::from_utf8(bytes).map_err(|_| format!("{what} is not valid UTF-8"))
}

// ---- Reading ---------------------------------------------------------------------

#[no_mangle]
pub unsafe extern "C" fn cxlsx_book_open(
    path: *const c_char,
    path_len: usize,
    error: *const cxlsx_error,
) -> *mut Book {
    guarded(error, std::ptr::null_mut(), || {
        // SAFETY: as documented in the header
        let path = unsafe { text_arg(path, path_len, "path") }?;
        Ok(Box::into_raw(Box::new(Book::open(Path::new(path))?)))
    })
}

#[no_mangle]
pub unsafe extern "C" fn cxlsx_book_close(book: *mut Book) {
    if !book.is_null() {
        // SAFETY: book comes from cxlsx_book_open and is closed once.
        let _ = catch_unwind(AssertUnwindSafe(|| drop(unsafe { Box::from_raw(book) })));
    }
}

#[no_mangle]
pub unsafe extern "C" fn cxlsx_book_sheet_count(book: *const Book) -> usize {
    // SAFETY: NULL or a book from cxlsx_book_open
    unsafe { book.as_ref() }.map_or(0, |b| b.sheet_names().len())
}

#[no_mangle]
pub unsafe extern "C" fn cxlsx_book_sheet_name(
    book: *const Book,
    index: usize,
    len: *mut usize,
) -> *const c_char {
    // SAFETY: NULL or a book from cxlsx_book_open
    let name = unsafe { book.as_ref() }.and_then(|b| b.sheet_names().get(index));
    // SAFETY: len is NULL or writable
    if let Some(len) = unsafe { len.as_mut() } {
        *len = name.map_or(0, String::len);
    }
    name.map_or(std::ptr::null(), |n| n.as_ptr().cast())
}

struct CVisitor<'a>(&'a cxlsx_sheet_visitor);

impl SheetVisitor for CVisitor<'_> {
    fn begin(&mut self, rows: u32, columns: u32) -> bool {
        self.0.begin.is_none_or(|f| f(self.0.ctx, rows, columns))
    }

    fn row(&mut self, row: u32, hidden: bool) -> bool {
        self.0.row.is_none_or(|f| f(self.0.ctx, row, hidden))
    }

    fn cell(&mut self, cell: &Cell) -> bool {
        let Some(f) = self.0.cell else { return true };
        let mut c = cxlsx_cell {
            row: cell.row,
            column: cell.column,
            kind: 0,
            boolean: false,
            number: 0.0,
            text: std::ptr::null(),
            text_len: 0,
        };
        match &cell.value {
            CellValue::Bool(b) => (c.kind, c.boolean) = (CXLSX_BOOL, *b),
            CellValue::Number(n) => (c.kind, c.number) = (CXLSX_NUMBER, *n),
            CellValue::DateTime(n) => (c.kind, c.number) = (CXLSX_DATETIME, *n),
            CellValue::Duration(n) => (c.kind, c.number) = (CXLSX_DURATION, *n),
            CellValue::Text(t) => {
                (c.kind, c.text, c.text_len) = (CXLSX_TEXT, t.as_ptr().cast(), t.len())
            }
            CellValue::Error(t) => {
                (c.kind, c.text, c.text_len) = (CXLSX_ERROR, t.as_ptr().cast(), t.len())
            }
            CellValue::UncachedFormula => c.kind = CXLSX_UNCACHED_FORMULA,
        }
        f(self.0.ctx, &c)
    }
}

#[no_mangle]
pub unsafe extern "C" fn cxlsx_book_read_sheet(
    book: *mut Book,
    name: *const c_char,
    name_len: usize,
    visitor: *const cxlsx_sheet_visitor,
    error: *const cxlsx_error,
) -> c_int {
    guarded(error, CXLSX_FAILED, || {
        // SAFETY: NULL or a book from cxlsx_book_open, used by one thread
        let book = unsafe { book.as_mut() }.ok_or("book is NULL")?;
        // SAFETY: as documented in the header
        let name = unsafe { text_arg(name, name_len, "sheet name") }?;
        // SAFETY: NULL or a valid visitor
        let visitor = unsafe { visitor.as_ref() }.ok_or("visitor is NULL")?;
        match book.read_sheet(name, &mut CVisitor(visitor)) {
            Ok(()) => Ok(CXLSX_OK),
            Err(ReadError::Stopped) => Ok(CXLSX_STOPPED),
            Err(ReadError::Failed(message)) => Err(message),
        }
    })
}

// ---- Writing ---------------------------------------------------------------------

#[no_mangle]
pub extern "C" fn cxlsx_writer_new() -> *mut Writer {
    catch_unwind(|| Box::into_raw(Box::new(Writer::new()))).unwrap_or(std::ptr::null_mut())
}

#[no_mangle]
pub unsafe extern "C" fn cxlsx_writer_free(writer: *mut Writer) {
    if !writer.is_null() {
        // SAFETY: writer comes from cxlsx_writer_new and is freed once.
        let _ = catch_unwind(AssertUnwindSafe(|| drop(unsafe { Box::from_raw(writer) })));
    }
}

/// SAFETY: writer is NULL or from cxlsx_writer_new, used by one thread.
unsafe fn writer_arg<'a>(writer: *mut Writer) -> Result<&'a mut Writer, String> {
    writer.as_mut().ok_or_else(|| "writer is NULL".to_string())
}

fn style_arg(style: u32) -> Result<Style, String> {
    Style::from_code(style).ok_or_else(|| format!("unknown style {style}"))
}

#[no_mangle]
pub unsafe extern "C" fn cxlsx_writer_add_sheet(
    writer: *mut Writer,
    name: *const c_char,
    name_len: usize,
    error: *const cxlsx_error,
) -> bool {
    guarded(error, false, || {
        // SAFETY: as documented in the header
        let (writer, name) =
            unsafe { (writer_arg(writer)?, text_arg(name, name_len, "sheet name")?) };
        writer.add_sheet(name).map(|_| true)
    })
}

#[no_mangle]
pub unsafe extern "C" fn cxlsx_writer_text(
    writer: *mut Writer,
    row: u32,
    column: u32,
    text: *const c_char,
    text_len: usize,
    style: u32,
    error: *const cxlsx_error,
) -> bool {
    guarded(error, false, || {
        // SAFETY: as documented in the header
        let (writer, text) = unsafe { (writer_arg(writer)?, text_arg(text, text_len, "text")?) };
        writer
            .text(row, column, text, style_arg(style)?)
            .map(|_| true)
    })
}

#[no_mangle]
pub unsafe extern "C" fn cxlsx_writer_number(
    writer: *mut Writer,
    row: u32,
    column: u32,
    value: f64,
    style: u32,
    error: *const cxlsx_error,
) -> bool {
    guarded(error, false, || {
        // SAFETY: as documented in the header
        let writer = unsafe { writer_arg(writer)? };
        writer
            .number(row, column, value, style_arg(style)?)
            .map(|_| true)
    })
}

#[no_mangle]
pub unsafe extern "C" fn cxlsx_writer_bool(
    writer: *mut Writer,
    row: u32,
    column: u32,
    value: bool,
    style: u32,
    error: *const cxlsx_error,
) -> bool {
    guarded(error, false, || {
        // SAFETY: as documented in the header
        let writer = unsafe { writer_arg(writer)? };
        writer
            .boolean(row, column, value, style_arg(style)?)
            .map(|_| true)
    })
}

#[no_mangle]
pub unsafe extern "C" fn cxlsx_writer_column_width(
    writer: *mut Writer,
    column: u32,
    width: f64,
    error: *const cxlsx_error,
) -> bool {
    guarded(error, false, || {
        // SAFETY: as documented in the header
        let writer = unsafe { writer_arg(writer)? };
        writer.column_width(column, width).map(|_| true)
    })
}

#[no_mangle]
pub unsafe extern "C" fn cxlsx_writer_save(
    writer: *mut Writer,
    path: *const c_char,
    path_len: usize,
    error: *const cxlsx_error,
) -> bool {
    guarded(error, false, || {
        // SAFETY: as documented in the header
        let (writer, path) = unsafe { (writer_arg(writer)?, text_arg(path, path_len, "path")?) };
        writer.save(Path::new(path)).map(|_| true)
    })
}
