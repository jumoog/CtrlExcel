//! Reading: calamine's streaming cell reader, merged with the row layout of
//! scan.rs. Never use `worksheet_range()`: it allocates rows x columns cells,
//! so a stray cell at XFD1048576 asks for 550 GB and aborts the process.

use crate::scan::{self, SheetLayout, WorkbookInfo};
use calamine::{DataRef, Reader, SheetType, Xlsx};
use std::borrow::Cow;
use std::io::Cursor;
use std::path::Path;
use std::sync::Arc;
use zip::ZipArchive;

/// Serial of 1904-01-01 in the 1900 date system.
const DATE1904_OFFSET: f64 = 1462.0;

/// The file's bytes, shared by calamine and the layout scan.
#[derive(Clone)]
struct Bytes(Arc<[u8]>);

impl AsRef<[u8]> for Bytes {
    fn as_ref(&self) -> &[u8] {
        &self.0
    }
}

#[derive(Debug, Clone, PartialEq)]
pub enum CellValue<'a> {
    Bool(bool),
    Number(f64),
    /// Serial in the 1900 date system, also for 1904 workbooks.
    DateTime(f64),
    /// Elapsed time ([h]:mm and the like), in days.
    Duration(f64),
    Text(Cow<'a, str>),
    /// Excel's error code, e.g. "#DIV/0!".
    Error(String),
    /// A formula whose result was never calculated (no cached value).
    UncachedFormula,
}

#[derive(Debug, Clone, PartialEq)]
pub struct Cell<'a> {
    /// 1-based
    pub row: u32,
    /// 1-based
    pub column: u32,
    pub value: CellValue<'a>,
}

/// Receives a sheet: begin() once, then row() for every row stored in the
/// sheet in ascending order, each followed by its cells (ascending columns;
/// cells without a value are left out). Returning false stops the read.
pub trait SheetVisitor {
    /// rows: number of the last row, columns: highest column (0, 0 if empty).
    fn begin(&mut self, rows: u32, columns: u32) -> bool;
    fn row(&mut self, row: u32, hidden: bool) -> bool;
    fn cell(&mut self, cell: &Cell) -> bool;
}

#[derive(Debug, PartialEq)]
pub enum ReadError {
    Failed(String),
    /// A visitor callback returned false.
    Stopped,
}

impl From<String> for ReadError {
    fn from(message: String) -> Self {
        ReadError::Failed(message)
    }
}

pub struct Book {
    xlsx: Xlsx<Cursor<Bytes>>,
    zip: ZipArchive<Cursor<Bytes>>,
    info: WorkbookInfo,
    worksheets: Vec<String>,
}

impl Book {
    pub fn open(path: &Path) -> Result<Book, String> {
        let data: Arc<[u8]> = std::fs::read(path).map_err(|e| e.to_string())?.into();
        let xlsx: Xlsx<_> =
            Xlsx::new(Cursor::new(Bytes(data.clone()))).map_err(|e| e.to_string())?;
        let mut zip = ZipArchive::new(Cursor::new(Bytes(data))).map_err(|e| e.to_string())?;
        let info = scan::workbook_info(&mut zip)?;
        // Worksheets only: chart and dialog sheets have no cells.
        let worksheets = xlsx
            .sheets_metadata()
            .iter()
            .filter(|s| matches!(s.typ, SheetType::WorkSheet))
            .map(|s| s.name.clone())
            .collect();
        Ok(Book {
            xlsx,
            zip,
            info,
            worksheets,
        })
    }

    /// Worksheet names in workbook order.
    pub fn sheet_names(&self) -> &[String] {
        &self.worksheets
    }

    pub fn read_sheet(
        &mut self,
        name: &str,
        visitor: &mut dyn SheetVisitor,
    ) -> Result<(), ReadError> {
        if !self.worksheets.iter().any(|w| w == name) {
            return Err(ReadError::Failed(format!("worksheet '{name}' not found")));
        }
        let path = self
            .info
            .sheet_path(name)
            .ok_or_else(|| format!("worksheet '{name}': part not found in the workbook"))?
            .to_string();
        let layout = scan::sheet_layout(&mut self.zip, &path)?;
        if !visitor.begin(layout.last_row(), layout.columns) {
            return Err(ReadError::Stopped);
        }
        let date1904 = self.info.date1904;
        let mut reader = self
            .xlsx
            .worksheet_cells_reader(name)
            .map_err(|e| e.to_string())?;
        let mut rows = RowCursor::new(&layout);
        while let Some(record) = reader
            .next_cell_with_formula_metadata()
            .map_err(|e| e.to_string())?
        {
            let Some(value) = cell_value(record.value, record.formula.is_some(), date1904) else {
                continue;
            };
            let cell = Cell {
                row: record.pos.0 + 1,
                column: record.pos.1 + 1,
                value,
            };
            if !rows.advance_to(cell.row, visitor) || !visitor.cell(&cell) {
                return Err(ReadError::Stopped);
            }
        }
        if !rows.finish(visitor) {
            return Err(ReadError::Stopped);
        }
        Ok(())
    }
}

/// Hands the layout's rows to the visitor in step with the cells.
struct RowCursor<'a> {
    rows: &'a [(u32, bool)],
    next: usize,
    current: u32, // row last handed out
}

impl<'a> RowCursor<'a> {
    fn new(layout: &'a SheetLayout) -> Self {
        RowCursor {
            rows: &layout.rows,
            next: 0,
            current: 0,
        }
    }

    /// Rows up to and including `row`.
    fn advance_to(&mut self, row: u32, visitor: &mut dyn SheetVisitor) -> bool {
        if row == self.current {
            return true;
        }
        while self.next < self.rows.len() && self.rows[self.next].0 <= row {
            let (number, hidden) = self.rows[self.next];
            self.next += 1;
            if !visitor.row(number, hidden) {
                return false;
            }
            self.current = number;
        }
        if self.current != row {
            // Not in the layout (rows out of order): calamine's position wins.
            self.current = row;
            return visitor.row(row, false);
        }
        true
    }

    fn finish(&mut self, visitor: &mut dyn SheetVisitor) -> bool {
        while self.next < self.rows.len() {
            let (number, hidden) = self.rows[self.next];
            self.next += 1;
            if !visitor.row(number, hidden) {
                return false;
            }
        }
        true
    }
}

/// None for a cell without a value.
fn cell_value(value: DataRef<'_>, has_formula: bool, date1904: bool) -> Option<CellValue<'_>> {
    Some(match value {
        DataRef::Empty if has_formula => CellValue::UncachedFormula,
        DataRef::Empty => return None,
        DataRef::Bool(b) => CellValue::Bool(b),
        DataRef::Int(i) => CellValue::Number(i as f64),
        DataRef::Float(f) => CellValue::Number(f),
        DataRef::DateTime(d) if d.is_duration() => CellValue::Duration(d.as_f64()),
        DataRef::DateTime(d) => CellValue::DateTime(if date1904 {
            d.as_f64() + DATE1904_OFFSET
        } else {
            d.as_f64()
        }),
        DataRef::SharedString(s) => CellValue::Text(without_nul(Cow::Borrowed(s))),
        DataRef::String(s) | DataRef::DateTimeIso(s) | DataRef::DurationIso(s) => {
            CellValue::Text(without_nul(Cow::Owned(s)))
        }
        DataRef::Error(e) => CellValue::Error(e.to_string()),
    })
}

/// U+0000 (the escape _x0000_) would end the C string and cut off the rest.
fn without_nul(text: Cow<'_, str>) -> Cow<'_, str> {
    if text.contains('\0') {
        Cow::Owned(text.replace('\0', ""))
    } else {
        text
    }
}
