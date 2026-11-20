//! What calamine does not expose, read from the package directly: the date
//! system and sheet paths of the workbook, and per sheet which rows exist,
//! which are hidden and how many columns the sheet spans.

use quick_xml::events::{BytesStart, Event};
use quick_xml::{Reader, XmlVersion};
use std::collections::HashMap;
use std::io::{BufReader, Read, Seek};
use zip::ZipArchive;

/// Excel's sheet limits.
pub const MAX_ROWS: u32 = 1_048_576;
pub const MAX_COLUMNS: u32 = 16_384;

pub struct WorkbookInfo {
    /// Serials count from 1904-01-01 (`<workbookPr date1904="1"/>`).
    pub date1904: bool,
    /// Sheet name -> path of its part in the package.
    sheet_paths: HashMap<String, String>,
}

impl WorkbookInfo {
    pub fn sheet_path(&self, name: &str) -> Option<&str> {
        self.sheet_paths.get(name).map(String::as_str)
    }
}

/// Rows stored in a sheet and the columns it spans, by the rules of
/// ECMA-376: a `<row>` or `<c>` without the optional `r` attribute follows
/// the previous one; one that does not ascend is skipped.
#[derive(Debug, Default, PartialEq)]
pub struct SheetLayout {
    /// 1-based row numbers in ascending order, with their hidden flag.
    pub rows: Vec<(u32, bool)>,
    /// Highest 1-based column of any `<c>` (also cells without a value).
    pub columns: u32,
}

impl SheetLayout {
    pub fn last_row(&self) -> u32 {
        self.rows.last().map_or(0, |r| r.0)
    }
}

fn local_name_is(e: &BytesStart, name: &str) -> bool {
    e.local_name().as_ref() == name
}

/// Raw value of the attribute with this local name (any prefix).
fn attribute(e: &BytesStart, name: &str) -> Result<Option<String>, String> {
    for attr in e.attributes().with_checks(false) {
        let attr = attr.map_err(|err| err.to_string())?;
        if attr.key.local_name().as_ref() == name {
            let value = attr
                .normalized_value(XmlVersion::Implicit1_0)
                .map_err(|err| err.to_string())?;
            return Ok(Some(value.into_owned()));
        }
    }
    Ok(None)
}

fn is_true(value: &str) -> bool {
    value == "1" || value.eq_ignore_ascii_case("true")
}

/// Column (1-based) of a cell reference such as "AB12"; 0 if it does not
/// start with a column of at most XFD.
pub fn reference_column(reference: &str) -> u32 {
    let mut column: u32 = 0;
    for b in reference.bytes() {
        if !b.is_ascii_uppercase() {
            break;
        }
        column = column * 26 + u32::from(b - b'A' + 1);
        if column > MAX_COLUMNS {
            return 0;
        }
    }
    column
}

fn read_part<R: Read + Seek>(
    zip: &mut ZipArchive<R>,
    path: &str,
) -> Result<Option<Vec<u8>>, String> {
    match zip.by_name(path) {
        Ok(mut file) => {
            let mut data = Vec::new();
            file.read_to_end(&mut data)
                .map_err(|e| format!("{path}: {e}"))?;
            Ok(Some(data))
        }
        Err(zip::result::ZipError::FileNotFound) => Ok(None),
        Err(e) => Err(format!("{path}: {e}")),
    }
}

/// Path of a relationship target, relative to the part in base_dir.
fn resolve_target(base_dir: &str, target: &str) -> String {
    if let Some(absolute) = target.strip_prefix('/') {
        return absolute.to_string();
    }
    let mut parts: Vec<&str> = base_dir.split('/').filter(|p| !p.is_empty()).collect();
    for segment in target.split('/') {
        match segment {
            "" | "." => {}
            ".." => {
                parts.pop();
            }
            s => parts.push(s),
        }
    }
    parts.join("/")
}

/// Id -> target of the relationships in a .rels part.
fn relationships(xml: &[u8], base_dir: &str) -> Result<HashMap<String, String>, String> {
    let mut reader = Reader::from_reader(xml);
    let mut buf = Vec::new();
    let mut rels = HashMap::new();
    loop {
        match reader
            .read_event_into(&mut buf)
            .map_err(|e| e.to_string())?
        {
            Event::Start(e) | Event::Empty(e) if local_name_is(&e, "Relationship") => {
                if let (Some(id), Some(target)) = (attribute(&e, "Id")?, attribute(&e, "Target")?) {
                    rels.insert(id, resolve_target(base_dir, &target));
                }
            }
            Event::Eof => break,
            _ => {}
        }
        buf.clear();
    }
    Ok(rels)
}

fn office_document_path<R: Read + Seek>(zip: &mut ZipArchive<R>) -> Result<String, String> {
    const OFFICE_DOCUMENT: &str = "/officeDocument";
    if let Some(xml) = read_part(zip, "_rels/.rels")? {
        let mut reader = Reader::from_reader(xml.as_slice());
        let mut buf = Vec::new();
        loop {
            match reader
                .read_event_into(&mut buf)
                .map_err(|e| e.to_string())?
            {
                Event::Start(e) | Event::Empty(e) if local_name_is(&e, "Relationship") => {
                    let is_document =
                        attribute(&e, "Type")?.is_some_and(|t| t.ends_with(OFFICE_DOCUMENT));
                    if is_document {
                        if let Some(target) = attribute(&e, "Target")? {
                            return Ok(resolve_target("", &target));
                        }
                    }
                }
                Event::Eof => break,
                _ => {}
            }
            buf.clear();
        }
    }
    Ok("xl/workbook.xml".to_string())
}

pub fn workbook_info<R: Read + Seek>(zip: &mut ZipArchive<R>) -> Result<WorkbookInfo, String> {
    let workbook_path = office_document_path(zip)?;
    let (dir, file) = workbook_path
        .rsplit_once('/')
        .unwrap_or(("", workbook_path.as_str()));
    let rels_path = if dir.is_empty() {
        format!("_rels/{file}.rels")
    } else {
        format!("{dir}/_rels/{file}.rels")
    };
    let rels = match read_part(zip, &rels_path)? {
        Some(xml) => relationships(&xml, dir)?,
        None => HashMap::new(),
    };
    let xml =
        read_part(zip, &workbook_path)?.ok_or_else(|| format!("{workbook_path} is missing"))?;

    let mut info = WorkbookInfo {
        date1904: false,
        sheet_paths: HashMap::new(),
    };
    let mut reader = Reader::from_reader(xml.as_slice());
    let mut buf = Vec::new();
    loop {
        match reader
            .read_event_into(&mut buf)
            .map_err(|e| format!("{workbook_path}: {e}"))?
        {
            Event::Start(e) | Event::Empty(e) => {
                if local_name_is(&e, "workbookPr") {
                    info.date1904 = attribute(&e, "date1904")?.is_some_and(|v| is_true(&v));
                } else if local_name_is(&e, "sheet") {
                    // r:id; sheetId has another local name
                    if let (Some(name), Some(id)) = (attribute(&e, "name")?, attribute(&e, "id")?) {
                        if let Some(path) = rels.get(&id) {
                            info.sheet_paths.insert(name, path.clone());
                        }
                    }
                }
            }
            Event::Eof => break,
            _ => {}
        }
        buf.clear();
    }
    Ok(info)
}

/// One pass over the sheet XML that only looks at `<row>` and `<c>` tags.
pub fn sheet_layout<R: Read + Seek>(
    zip: &mut ZipArchive<R>,
    path: &str,
) -> Result<SheetLayout, String> {
    let file = zip.by_name(path).map_err(|e| format!("{path}: {e}"))?;
    let mut reader = Reader::from_reader(BufReader::with_capacity(64 * 1024, file));
    let mut buf = Vec::new();
    let mut layout = SheetLayout::default();
    let mut in_sheet_data = false;
    let mut in_row = false; // inside a <row> that is kept
    let mut row: u32 = 0; // number of the previous <row>
    let mut column: u32 = 0; // column of the previous <c> in the row
    loop {
        let event = reader
            .read_event_into(&mut buf)
            .map_err(|e| format!("{path}: {e}"))?;
        match event {
            Event::Start(ref e) | Event::Empty(ref e) => {
                let empty = matches!(event, Event::Empty(_));
                if local_name_is(e, "sheetData") {
                    in_sheet_data = !empty;
                } else if in_sheet_data && local_name_is(e, "row") {
                    let r = attribute(e, "r")?
                        .and_then(|v| v.parse::<u32>().ok())
                        .unwrap_or(0);
                    let number = if r == 0 { row + 1 } else { r };
                    in_row = number > row && number <= MAX_ROWS;
                    if in_row {
                        row = number;
                        column = 0;
                        let hidden = attribute(e, "hidden")?.is_some_and(|v| is_true(&v));
                        layout.rows.push((number, hidden));
                    }
                    if empty {
                        in_row = false;
                    }
                } else if in_row && local_name_is(e, "c") {
                    let mut c = attribute(e, "r")?.as_deref().map_or(0, reference_column);
                    if c == 0 {
                        c = column + 1;
                    }
                    if c > column && c <= MAX_COLUMNS {
                        column = c;
                        layout.columns = layout.columns.max(c);
                    }
                }
            }
            Event::End(ref e) => match e.local_name().as_ref() {
                "sheetData" => break,
                "row" => in_row = false,
                _ => {}
            },
            Event::Eof => break,
            _ => {}
        }
        buf.clear();
    }
    Ok(layout)
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn columns() {
        assert_eq!(reference_column("A1"), 1);
        assert_eq!(reference_column("Z9"), 26);
        assert_eq!(reference_column("AA1"), 27);
        assert_eq!(reference_column("XFD1048576"), 16_384);
        assert_eq!(reference_column("XFE1"), 0);
        assert_eq!(reference_column("1"), 0);
        assert_eq!(reference_column(""), 0);
        assert_eq!(reference_column("a1"), 0);
    }

    #[test]
    fn targets() {
        assert_eq!(
            resolve_target("xl", "worksheets/sheet1.xml"),
            "xl/worksheets/sheet1.xml"
        );
        assert_eq!(
            resolve_target("xl", "/xl/worksheets/sheet1.xml"),
            "xl/worksheets/sheet1.xml"
        );
        assert_eq!(
            resolve_target("xl/sub", "../worksheets/s.xml"),
            "xl/worksheets/s.xml"
        );
        assert_eq!(resolve_target("", "xl/workbook.xml"), "xl/workbook.xml");
    }
}
