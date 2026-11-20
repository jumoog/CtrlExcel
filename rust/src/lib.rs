//! .xlsx reading (calamine) and writing (rust_xlsxwriter) for the CTRL
//! extension CtrlExcelReader, which links this crate as a static library
//! through the C ABI in include/ctrlexcel_xlsx.h (ffi.rs). Converting values
//! to and from WinCC OA types (dates in local time, int/long/float, header
//! keys) stays in C++; this crate only moves typed cells in and out of files.

#![allow(non_camel_case_types)]

pub mod ffi;
pub mod reader;
pub mod scan;
pub mod writer;

pub use reader::{Book, Cell, CellValue, ReadError, SheetVisitor};
pub use writer::{Style, Writer};
