//! Writing with rust_xlsxwriter. The file is written to a temporary file next
//! to the target and renamed over it, so a failed write never leaves a
//! partial file behind, and an existing file is only replaced once the new
//! one is complete. A symbolic link is followed (the file it points to is
//! replaced), and the replaced file's permissions are kept on Unix.

use rust_xlsxwriter::{DocProperties, Format, Formula, Workbook, Worksheet, XlsxError};
use std::path::{Path, PathBuf};
use std::sync::atomic::{AtomicU64, Ordering};

/// Excel's maximum column width.
pub const MAX_COLUMN_WIDTH: f64 = 255.0;

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum Style {
    Default,
    /// Bold, for the header row.
    Header,
    /// Built-in number format 22, "m/d/yyyy h:mm".
    DateTime,
}

impl Style {
    pub fn from_code(code: u32) -> Option<Style> {
        match code {
            0 => Some(Style::Default),
            1 => Some(Style::Header),
            2 => Some(Style::DateTime),
            _ => None,
        }
    }
}

pub struct Writer {
    book: Workbook,
    sheet: Option<usize>,
    sheets: usize,
    header: Format,
    datetime: Format,
}

fn message(e: XlsxError) -> String {
    e.to_string()
}

/// rust_xlsxwriter counts rows and columns from 0.
fn position(row: u32, column: u32) -> Result<(u32, u16), String> {
    if row == 0 || column == 0 || column > u32::from(u16::MAX) {
        return Err(format!("invalid cell position row {row}, column {column}"));
    }
    Ok((row - 1, (column - 1) as u16))
}

impl Default for Writer {
    fn default() -> Self {
        Self::new()
    }
}

impl Writer {
    pub fn new() -> Writer {
        let mut book = Workbook::new();
        book.set_properties(&DocProperties::new().set_author("WinCC OA"));
        Writer {
            book,
            sheet: None,
            sheets: 0,
            header: Format::new().set_bold(),
            datetime: Format::new().set_num_format_index(22),
        }
    }

    /// Adds a worksheet; the cell functions write to the last one added.
    pub fn add_sheet(&mut self, name: &str) -> Result<(), String> {
        // Low memory: finished rows go to a temporary file, shared strings
        // stay (smaller files than constant memory's inline strings).
        let sheet = self.book.add_worksheet_with_low_memory().map_err(message)?;
        sheet.set_name(name).map_err(message)?;
        self.sheet = Some(self.sheets);
        self.sheets += 1;
        Ok(())
    }

    fn current(&mut self) -> Result<(&mut Worksheet, &Format, &Format), String> {
        let index = self.sheet.ok_or_else(|| "no worksheet added".to_string())?;
        let sheet = self.book.worksheet_from_index(index).map_err(message)?;
        Ok((sheet, &self.header, &self.datetime))
    }

    fn format<'f>(style: Style, header: &'f Format, datetime: &'f Format) -> Option<&'f Format> {
        match style {
            Style::Default => None,
            Style::Header => Some(header),
            Style::DateTime => Some(datetime),
        }
    }

    pub fn text(&mut self, row: u32, column: u32, text: &str, style: Style) -> Result<(), String> {
        let (r, c) = position(row, column)?;
        let (sheet, header, datetime) = self.current()?;
        match Self::format(style, header, datetime) {
            Some(format) => sheet.write_string_with_format(r, c, text, format),
            None => sheet.write_string(r, c, text),
        }
        .map(|_| ())
        .map_err(message)
    }

    /// Non-finite values become the error #NUM!, as Excel shows for them.
    pub fn number(
        &mut self,
        row: u32,
        column: u32,
        value: f64,
        style: Style,
    ) -> Result<(), String> {
        let (r, c) = position(row, column)?;
        let (sheet, header, datetime) = self.current()?;
        if !value.is_finite() {
            return sheet
                .write_formula(r, c, Formula::new("=#NUM!").set_result("#NUM!"))
                .map(|_| ())
                .map_err(message);
        }
        match Self::format(style, header, datetime) {
            Some(format) => sheet.write_number_with_format(r, c, value, format),
            None => sheet.write_number(r, c, value),
        }
        .map(|_| ())
        .map_err(message)
    }

    pub fn boolean(
        &mut self,
        row: u32,
        column: u32,
        value: bool,
        style: Style,
    ) -> Result<(), String> {
        let (r, c) = position(row, column)?;
        let (sheet, header, datetime) = self.current()?;
        match Self::format(style, header, datetime) {
            Some(format) => sheet.write_boolean_with_format(r, c, value, format),
            None => sheet.write_boolean(r, c, value),
        }
        .map(|_| ())
        .map_err(message)
    }

    /// Width in characters, capped at Excel's maximum.
    pub fn column_width(&mut self, column: u32, width: f64) -> Result<(), String> {
        let (_, c) = position(1, column)?;
        let (sheet, _, _) = self.current()?;
        sheet
            .set_column_width(c, width.clamp(0.0, MAX_COLUMN_WIDTH))
            .map(|_| ())
            .map_err(message)
    }

    /// Writes the workbook to path (replacing an existing file). A workbook
    /// without sheets gets an empty "Sheet1", as Excel requires one.
    pub fn save(&mut self, path: &Path) -> Result<(), String> {
        if self.sheets == 0 {
            self.add_sheet("Sheet1")?;
        }
        let target = rename_target(path);
        let temp = temp_path(&target);
        let result = self
            .book
            .save(&temp)
            .map_err(message)
            .and_then(|_| keep_permissions(&target, &temp))
            .and_then(|_| std::fs::rename(&temp, &target).map_err(|e| e.to_string()));
        if result.is_err() {
            let _ = std::fs::remove_file(&temp);
        }
        result
    }
}

/// The file to replace: for a symbolic link the file it points to (also one
/// that does not exist yet), so the rename writes through the link instead of
/// replacing it.
fn rename_target(path: &Path) -> PathBuf {
    let mut target = path.to_path_buf();
    // A bounded number of hops, as the OS allows: a link cycle stops here.
    for _ in 0..40 {
        let is_link = target
            .symlink_metadata()
            .is_ok_and(|m| m.file_type().is_symlink());
        if !is_link {
            break;
        }
        let Ok(link) = std::fs::read_link(&target) else {
            break;
        };
        // A relative link is relative to the directory that holds it.
        target = match target.parent() {
            Some(dir) if link.is_relative() => dir.join(link),
            _ => link,
        };
    }
    target
}

/// Gives the new file the permissions of the one it replaces (a new file gets
/// the process's defaults). Owner and ACLs are not carried over.
fn keep_permissions(target: &Path, temp: &Path) -> Result<(), String> {
    match std::fs::metadata(target) {
        // On Windows the permissions are only the read-only flag, which would
        // make the rename fail; the caller rejects read-only targets anyway.
        Ok(meta) if cfg!(unix) => {
            std::fs::set_permissions(temp, meta.permissions()).map_err(|e| e.to_string())
        }
        _ => Ok(()),
    }
}

/// "<name>.~<pid>-<n>.tmp" in the target's directory, so the rename stays on
/// one file system (and is atomic there).
fn temp_path(path: &Path) -> PathBuf {
    static COUNTER: AtomicU64 = AtomicU64::new(0);
    let n = COUNTER.fetch_add(1, Ordering::Relaxed);
    let mut name = path
        .file_name()
        .map(|n| n.to_os_string())
        .unwrap_or_default();
    name.push(format!(".~{}-{n}.tmp", std::process::id()));
    path.with_file_name(name)
}
