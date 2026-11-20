//! What calamine does not expose, read from the package directly: the date
//! system and sheet paths of the workbook, and per sheet which rows exist,
//! which are hidden and how many columns the sheet spans.

use quick_xml::events::{BytesStart, Event};
use quick_xml::{Reader, XmlVersion};
use std::collections::HashMap;
use std::io::{Read, Seek};
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
fn reference_column(reference: &[u8]) -> u32 {
    let mut column: u32 = 0;
    for &b in reference {
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

/// Rows and columns of a sheet, by the rules in SheetLayout, from one pass
/// over the sheet XML. It only needs the `r` and `hidden` attributes of
/// `<row>` and `<c>` tags, so it jumps from tag to tag (memchr) instead of
/// parsing the XML, which matters because the scan runs in addition to
/// calamine's own pass. The XML is read in chunks, so memory does not grow
/// with the sheet.
pub fn sheet_layout<R: Read + Seek>(
    zip: &mut ZipArchive<R>,
    path: &str,
) -> Result<SheetLayout, String> {
    let mut file = zip.by_name(path).map_err(|e| format!("{path}: {e}"))?;
    let mut scanner = LayoutScanner::default();
    let mut buf = vec![0u8; 256 * 1024];
    let mut len = 0; // bytes in buf not yet scanned
    loop {
        if len == buf.len() {
            // A single tag (or comment) longer than the buffer.
            buf.resize(buf.len() * 2, 0);
        }
        let n = file
            .read(&mut buf[len..])
            .map_err(|e| format!("{path}: {e}"))?;
        len += n;
        let consumed = scanner.scan(&buf[..len]);
        if scanner.done || n == 0 {
            break;
        }
        buf.copy_within(consumed..len, 0);
        len -= consumed;
    }
    Ok(scanner.layout)
}

#[derive(Default)]
struct LayoutScanner {
    layout: SheetLayout,
    in_sheet_data: bool,
    in_row: bool, // inside a <row> that is kept
    row: u32,     // number of the previous <row>
    column: u32,  // column of the previous <c> in the row
    done: bool,   // </sheetData> seen
}

impl LayoutScanner {
    /// Scans the complete tags in data; returns how many bytes were used (the
    /// rest starts with an incomplete tag and is passed again with more data).
    fn scan(&mut self, data: &[u8]) -> usize {
        let mut pos = 0;
        while !self.done {
            let Some(lt) = memchr::memchr(b'<', &data[pos..]) else {
                return data.len(); // text only
            };
            let start = pos + lt;
            let rest = &data[start + 1..];
            // Comments and CDATA sections may contain '<' and '>'.
            let end = if rest.starts_with(b"!--") {
                memchr::memmem::find(rest, b"-->").map(|i| start + 1 + i + 3)
            } else if rest.starts_with(b"![CDATA[") {
                memchr::memmem::find(rest, b"]]>").map(|i| start + 1 + i + 3)
            } else {
                tag_end(rest).map(|i| {
                    self.tag(&rest[..i]);
                    start + 1 + i + 1
                })
            };
            match end {
                Some(end) => pos = end,
                None => return start,
            }
        }
        data.len()
    }

    /// tag: the bytes between '<' and '>'.
    fn tag(&mut self, tag: &[u8]) {
        if let Some(name) = tag.strip_prefix(b"/") {
            match local_name(name.trim_ascii_end()) {
                b"sheetData" => self.done = self.in_sheet_data,
                b"row" => self.in_row = false,
                _ => {}
            }
            return;
        }
        let self_closing = tag.ends_with(b"/");
        let name_len = tag
            .iter()
            .position(|b| b.is_ascii_whitespace() || *b == b'/')
            .unwrap_or(tag.len());
        let attributes = &tag[name_len..];
        match local_name(&tag[..name_len]) {
            b"sheetData" => self.in_sheet_data = !self_closing,
            b"row" if self.in_sheet_data => {
                let r = tag_attribute(attributes, b"r").map_or(0, parse_u32);
                let number = if r == 0 { self.row + 1 } else { r };
                self.in_row = number > self.row && number <= MAX_ROWS;
                if self.in_row {
                    self.row = number;
                    self.column = 0;
                    let hidden = tag_attribute(attributes, b"hidden")
                        .is_some_and(|v| v == b"1" || v.eq_ignore_ascii_case(b"true"));
                    self.layout.rows.push((number, hidden));
                }
                if self_closing {
                    self.in_row = false;
                }
            }
            b"c" if self.in_row => {
                let mut c = tag_attribute(attributes, b"r").map_or(0, reference_column);
                if c == 0 {
                    c = self.column + 1;
                }
                if c > self.column && c <= MAX_COLUMNS {
                    self.column = c;
                    self.layout.columns = self.layout.columns.max(c);
                }
            }
            _ => {}
        }
    }
}

/// Index of the '>' that ends the tag whose bytes start at data[0], skipping
/// quoted attribute values (which may contain '>').
fn tag_end(data: &[u8]) -> Option<usize> {
    let mut pos = 0;
    loop {
        let i = pos + memchr::memchr3(b'>', b'"', b'\'', &data[pos..])?;
        if data[i] == b'>' {
            return Some(i);
        }
        let close = memchr::memchr(data[i], &data[i + 1..])?;
        pos = i + 1 + close + 1;
    }
}

/// The name without a namespace prefix ("x:row" -> "row").
fn local_name(name: &[u8]) -> &[u8] {
    match memchr::memchr(b':', name) {
        Some(i) => &name[i + 1..],
        None => name,
    }
}

/// Raw value of the attribute with this local name in the attribute part of
/// a tag. No entity decoding: the attributes read here never need it.
fn tag_attribute<'a>(mut attributes: &'a [u8], wanted: &[u8]) -> Option<&'a [u8]> {
    loop {
        let start = attributes.iter().position(|b| !b.is_ascii_whitespace())?;
        attributes = &attributes[start..];
        let eq = memchr::memchr(b'=', attributes)?;
        let name = attributes[..eq].trim_ascii_end();
        let after = attributes[eq + 1..].trim_ascii_start();
        let quote = *after.first()?;
        if quote != b'"' && quote != b'\'' {
            return None;
        }
        let len = memchr::memchr(quote, &after[1..])?;
        if local_name(name) == wanted {
            return Some(&after[1..1 + len]);
        }
        attributes = &after[1 + len + 1..];
    }
}

fn parse_u32(digits: &[u8]) -> u32 {
    std::str::from_utf8(digits)
        .ok()
        .and_then(|s| s.parse().ok())
        .unwrap_or(0)
}

#[cfg(test)]
mod tests {
    use super::*;

    /// The scanner over data in chunks of `chunk` bytes, as sheet_layout feeds it.
    fn scan_in_chunks(data: &[u8], chunk: usize) -> SheetLayout {
        let mut scanner = LayoutScanner::default();
        let mut pending: Vec<u8> = Vec::new();
        for piece in data.chunks(chunk) {
            pending.extend_from_slice(piece);
            let used = scanner.scan(&pending);
            pending.drain(..used);
            if scanner.done {
                break;
            }
        }
        scanner.layout
    }

    #[test]
    fn layout_scanner() {
        let xml = br#"<?xml version="1.0"?><x:worksheet xmlns:x="ns"><x:sheetPr codeName="a>b"/>
<!-- <row r="99"><c r="ZZ99"/></row> --><x:sheetData>
<x:row r="2" spans="1:3" hidden="1"><x:c r="A2" t="s"><x:v>0</x:v></x:c><x:c r='C2' s="1"/></x:row>
<row hidden="true"><c><v>1</v></c><c><v><![CDATA[<c r="XFD3">]]></v></c></row>
<row r="2"><c r="E2"/></row>
<row r = "7" /><row><c r="B8"/><c r="A8"/><extLst/></row>
</x:sheetData><x:row r="100"><x:c r="Z100"/></x:row></x:worksheet>"#;
        let expected = SheetLayout {
            // row 2 again does not ascend; the row after 7 is 8; rows after
            // </sheetData> do not count
            rows: vec![(2, true), (3, true), (7, false), (8, false)],
            // C2; the CDATA and comment are no cells; A8 does not ascend
            columns: 3,
        };
        for chunk in [1, 2, 3, 7, 64, xml.len()] {
            assert_eq!(scan_in_chunks(xml, chunk), expected, "chunk size {chunk}");
        }
    }

    #[test]
    fn tag_attributes() {
        assert_eq!(tag_attribute(b" r=\"A1\" s=\"2\"", b"s"), Some(&b"2"[..]));
        assert_eq!(tag_attribute(b" x:r = 'B2'", b"r"), Some(&b"B2"[..]));
        assert_eq!(tag_attribute(b" spans=\"1:3\"", b"r"), None);
        assert_eq!(tag_attribute(b"", b"r"), None);
        assert_eq!(tag_end(b"c r=\"a>b\" s='>'>rest"), Some(15));
        assert_eq!(tag_end(b"c r=\"open"), None);
    }

    #[test]
    fn columns() {
        assert_eq!(reference_column(b"A1"), 1);
        assert_eq!(reference_column(b"Z9"), 26);
        assert_eq!(reference_column(b"AA1"), 27);
        assert_eq!(reference_column(b"XFD1048576"), 16_384);
        assert_eq!(reference_column(b"XFE1"), 0);
        assert_eq!(reference_column(b"1"), 0);
        assert_eq!(reference_column(b""), 0);
        assert_eq!(reference_column(b"a1"), 0);
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
