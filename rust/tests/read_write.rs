//! Reading the fixture (testdata/make_fixture.py) and write -> read round trips.

use ctrlexcel_xlsx::{Book, Cell, CellValue, ReadError, SheetVisitor, Style, Writer};
use std::path::{Path, PathBuf};

fn fixture() -> PathBuf {
    PathBuf::from(env!("CARGO_MANIFEST_DIR")).join("../testdata/CtrlExcelReaderFixture.xlsx")
}

fn temp(name: &str) -> PathBuf {
    std::env::temp_dir().join(format!(
        "ctrlexcel_xlsx_test_{}_{name}.xlsx",
        std::process::id()
    ))
}

#[derive(Debug, Clone, PartialEq)]
enum Event {
    Begin(u32, u32),
    Row(u32, bool),
    Cell(u32, u32, CellValue<'static>),
}

#[derive(Default)]
struct Recorder {
    events: Vec<Event>,
    stop_after: Option<usize>,
}

impl Recorder {
    fn push(&mut self, e: Event) -> bool {
        self.events.push(e);
        self.stop_after.is_none_or(|n| self.events.len() < n)
    }

    fn cells(&self) -> Vec<(u32, u32, CellValue<'static>)> {
        self.events
            .iter()
            .filter_map(|e| match e {
                Event::Cell(r, c, v) => Some((*r, *c, v.clone())),
                _ => None,
            })
            .collect()
    }

    fn value(&self, row: u32, column: u32) -> Option<CellValue<'static>> {
        self.cells()
            .into_iter()
            .find(|(r, c, _)| *r == row && *c == column)
            .map(|(_, _, v)| v)
    }
}

fn owned(v: &CellValue) -> CellValue<'static> {
    match v {
        CellValue::Text(t) => CellValue::Text(t.to_string().into()),
        CellValue::Bool(b) => CellValue::Bool(*b),
        CellValue::Number(n) => CellValue::Number(*n),
        CellValue::DateTime(n) => CellValue::DateTime(*n),
        CellValue::Duration(n) => CellValue::Duration(*n),
        CellValue::Error(e) => CellValue::Error(e.clone()),
        CellValue::UncachedFormula => CellValue::UncachedFormula,
    }
}

impl SheetVisitor for Recorder {
    fn begin(&mut self, rows: u32, columns: u32) -> bool {
        self.push(Event::Begin(rows, columns))
    }
    fn row(&mut self, row: u32, hidden: bool) -> bool {
        self.push(Event::Row(row, hidden))
    }
    fn cell(&mut self, cell: &Cell) -> bool {
        self.push(Event::Cell(cell.row, cell.column, owned(&cell.value)))
    }
}

fn read(path: &Path, sheet: &str) -> Recorder {
    let mut book = Book::open(path).unwrap();
    let mut rec = Recorder::default();
    book.read_sheet(sheet, &mut rec).unwrap();
    rec
}

fn text(t: &str) -> CellValue<'static> {
    CellValue::Text(t.to_string().into())
}

#[test]
fn fixture_sheet_names() {
    let book = Book::open(&fixture()).unwrap();
    assert_eq!(book.sheet_names(), ["Fixture", "Implicit"]);
}

#[test]
fn fixture_values() {
    let rec = read(&fixture(), "Fixture");
    assert_eq!(rec.events[0], Event::Begin(12, 2));
    let rows: Vec<u32> = rec
        .events
        .iter()
        .filter_map(|e| {
            if let Event::Row(r, _) = e {
                Some(*r)
            } else {
                None
            }
        })
        .collect();
    assert_eq!(
        rows,
        (1..=12).collect::<Vec<_>>(),
        "row 11 only carries formatting but exists"
    );
    assert_eq!(rec.value(2, 1), Some(text("pre1970")));
    assert_eq!(rec.value(2, 2), Some(CellValue::DateTime(22037.0)));
    assert_eq!(rec.value(3, 2), Some(CellValue::DateTime(46023.0)));
    assert_eq!(rec.value(4, 2), Some(CellValue::Duration(1.5)));
    assert_eq!(rec.value(5, 2), Some(CellValue::Error("#DIV/0!".into())));
    assert_eq!(rec.value(6, 2), Some(CellValue::UncachedFormula));
    assert_eq!(rec.value(7, 2), Some(text("a\r\nb")));
    // calamine decodes every _xHHHH_ escape, as ECMA-376 specifies
    assert_eq!(rec.value(8, 2), Some(text("ABA7")));
    assert_eq!(rec.value(9, 2), Some(CellValue::Number(1e20)));
    assert_eq!(
        rec.value(10, 2),
        Some(CellValue::Number(12345678901234567890.0))
    );
    assert_eq!(rec.value(11, 1), None);
    assert_eq!(rec.value(12, 2), Some(text("last")));
}

#[test]
fn fixture_implicit_references() {
    let rec = read(&fixture(), "Implicit");
    let rows: Vec<Event> = rec
        .events
        .iter()
        .filter(|e| !matches!(e, Event::Cell(..)))
        .cloned()
        .collect();
    assert_eq!(
        rows,
        [
            Event::Begin(6, 2),
            Event::Row(1, false),
            Event::Row(2, false),
            Event::Row(4, false),
            Event::Row(5, false),
            Event::Row(6, false)
        ]
    );
    assert_eq!(rec.value(5, 1), Some(text("implicitRow")));
    assert_eq!(rec.value(5, 2), Some(CellValue::Number(3.0)));
    assert_eq!(rec.value(6, 1), Some(text("implicitCell")));
    assert_eq!(rec.value(6, 2), Some(CellValue::Number(4.0)));
}

#[test]
fn events_are_ordered() {
    // Every cell follows the row() of its row, rows ascend.
    let rec = read(&fixture(), "Fixture");
    let mut row = 0;
    for e in &rec.events[1..] {
        match e {
            Event::Row(r, _) => {
                assert!(*r > row);
                row = *r;
            }
            Event::Cell(r, _, _) => assert_eq!(*r, row),
            Event::Begin(..) => panic!("second begin"),
        }
    }
}

#[test]
fn visitor_can_stop() {
    let mut book = Book::open(&fixture()).unwrap();
    let mut rec = Recorder {
        stop_after: Some(3),
        ..Default::default()
    };
    assert_eq!(
        book.read_sheet("Fixture", &mut rec),
        Err(ReadError::Stopped)
    );
    assert_eq!(rec.events.len(), 3);
}

#[test]
fn missing_sheet_and_file() {
    let mut book = Book::open(&fixture()).unwrap();
    let mut rec = Recorder::default();
    assert!(matches!(
        book.read_sheet("Nope", &mut rec),
        Err(ReadError::Failed(_))
    ));
    assert!(Book::open(&PathBuf::from("does/not/exist.xlsx")).is_err());
}

#[test]
fn round_trip() {
    let path = temp("round_trip");
    let mut w = Writer::new();
    w.add_sheet("Data").unwrap();
    w.text(1, 1, "Name", Style::Header).unwrap();
    w.text(1, 2, "When", Style::Header).unwrap();
    w.text(1, 3, "Flag", Style::Header).unwrap();
    w.text(2, 1, "a\r\nb\u{1}_x0041_\u{FFFE}\u{1F600}", Style::Default)
        .unwrap();
    w.number(2, 2, 46023.5, Style::DateTime).unwrap();
    w.boolean(2, 3, true, Style::Default).unwrap();
    w.number(3, 1, 3_000_000_000.0, Style::Default).unwrap();
    w.number(3, 2, f64::NAN, Style::Default).unwrap();
    w.number(3, 3, -0.25, Style::Default).unwrap();
    w.column_width(1, 1000.0).unwrap();
    w.add_sheet("Second").unwrap();
    w.text(1, 1, "x", Style::Default).unwrap();
    w.save(&path).unwrap();

    let book = Book::open(&path).unwrap();
    assert_eq!(book.sheet_names(), ["Data", "Second"]);
    let rec = read(&path, "Data");
    assert_eq!(rec.events[0], Event::Begin(3, 3));
    assert_eq!(rec.value(1, 2), Some(text("When")));
    assert_eq!(
        rec.value(2, 1),
        Some(text("a\r\nb\u{1}_x0041_\u{FFFE}\u{1F600}"))
    );
    assert_eq!(rec.value(2, 2), Some(CellValue::DateTime(46023.5)));
    assert_eq!(rec.value(2, 3), Some(CellValue::Bool(true)));
    assert_eq!(rec.value(3, 1), Some(CellValue::Number(3_000_000_000.0)));
    assert_eq!(rec.value(3, 2), Some(CellValue::Error("#NUM!".into())));
    assert_eq!(rec.value(3, 3), Some(CellValue::Number(-0.25)));
    assert_eq!(read(&path, "Second").value(1, 1), Some(text("x")));
    std::fs::remove_file(&path).unwrap();
}

#[test]
fn save_replaces_and_leaves_no_temp_file() {
    let path = temp("replace");
    std::fs::write(&path, b"old").unwrap();
    let mut w = Writer::new();
    w.save(&path).unwrap();
    let book = Book::open(&path).unwrap();
    assert_eq!(book.sheet_names(), ["Sheet1"]);
    let dir = path.parent().unwrap();
    let stem = path.file_name().unwrap().to_string_lossy().to_string();
    let leftovers: Vec<_> = std::fs::read_dir(dir)
        .unwrap()
        .filter_map(|e| e.ok())
        .filter(|e| {
            let n = e.file_name().to_string_lossy().to_string();
            n.starts_with(&stem) && n != stem
        })
        .collect();
    assert!(leftovers.is_empty(), "{leftovers:?}");
    std::fs::remove_file(&path).unwrap();
}

#[test]
fn failed_save_keeps_existing_file() {
    // A directory of that name cannot be replaced by the rename.
    let path = temp("dir_target");
    std::fs::create_dir_all(&path).unwrap();
    let mut w = Writer::new();
    w.add_sheet("S").unwrap();
    assert!(w.save(&path).is_err());
    assert!(path.is_dir());
    std::fs::remove_dir(&path).unwrap();
}

#[test]
fn string_limit_counts_utf16() {
    let mut w = Writer::new();
    w.add_sheet("S").unwrap();
    assert!(w.text(1, 1, &"a".repeat(32767), Style::Default).is_ok());
    assert!(w
        .text(2, 1, &"\u{1F600}".repeat(16384), Style::Default)
        .is_err());
}

#[test]
fn hidden_rows_and_stray_cell() {
    let path = temp("hidden");
    let mut book = rust_xlsxwriter::Workbook::new();
    let sheet = book.add_worksheet();
    sheet.write_string(0, 0, "Label").unwrap();
    sheet.write_string(1, 0, "visible").unwrap();
    sheet.write_string(2, 0, "hidden").unwrap();
    sheet.set_row_hidden(2).unwrap();
    sheet.set_row_hidden(5).unwrap(); // a hidden row without cells
    sheet.write_string(1_048_575, 16_383, "stray").unwrap(); // XFD1048576
    book.save(&path).unwrap();

    let rec = read(&path, "Sheet1");
    let rows: Vec<Event> = rec
        .events
        .iter()
        .filter(|e| !matches!(e, Event::Cell(..)))
        .cloned()
        .collect();
    assert_eq!(
        rows,
        [
            Event::Begin(1_048_576, 16_384),
            Event::Row(1, false),
            Event::Row(2, false),
            Event::Row(3, true),
            Event::Row(6, true),
            Event::Row(1_048_576, false)
        ]
    );
    assert_eq!(rec.value(1_048_576, 16_384), Some(text("stray")));
    std::fs::remove_file(&path).unwrap();
}

/// Copy of the package at path with xl/workbook.xml changed by edit.
fn patch_workbook_xml(path: &Path, out: &Path, edit: impl Fn(&str) -> String) {
    use std::io::{Read, Write};
    let mut zip = zip::ZipArchive::new(std::fs::File::open(path).unwrap()).unwrap();
    let mut writer = zip::ZipWriter::new(std::fs::File::create(out).unwrap());
    for i in 0..zip.len() {
        let mut entry = zip.by_index(i).unwrap();
        let name = entry.name().unwrap().to_string();
        let mut data = Vec::new();
        entry.read_to_end(&mut data).unwrap();
        if name == "xl/workbook.xml" {
            data = edit(std::str::from_utf8(&data).unwrap()).into_bytes();
        }
        writer
            .start_file(
                name,
                zip::write::SimpleFileOptions::default()
                    .compression_method(zip::CompressionMethod::Deflated),
            )
            .unwrap();
        writer.write_all(&data).unwrap();
    }
    writer.finish().unwrap();
}

#[test]
fn date1904_serials_are_converted() {
    let written = temp("date1904_src");
    let patched = temp("date1904");
    let mut w = Writer::new();
    w.add_sheet("S").unwrap();
    // 2026-01-01 is serial 44561 in the 1904 system, 46023 in the 1900 one.
    w.number(1, 1, 44561.5, Style::DateTime).unwrap();
    w.number(2, 1, 1.5, Style::Default).unwrap();
    w.save(&written).unwrap();
    patch_workbook_xml(&written, &patched, |xml| {
        if xml.contains("<workbookPr") {
            xml.replacen("<workbookPr", "<workbookPr date1904=\"1\"", 1)
        } else {
            let start = xml.find("<workbook ").unwrap();
            let end = start + xml[start..].find('>').unwrap() + 1;
            format!("{}<workbookPr date1904=\"1\"/>{}", &xml[..end], &xml[end..])
        }
    });
    let rec = read(&patched, "S");
    assert_eq!(rec.value(1, 1), Some(CellValue::DateTime(46023.5)));
    assert_eq!(
        rec.value(2, 1),
        Some(CellValue::Number(1.5)),
        "plain numbers stay as they are"
    );
    assert_eq!(
        read(&written, "S").value(1, 1),
        Some(CellValue::DateTime(44561.5))
    );
    std::fs::remove_file(&written).unwrap();
    std::fs::remove_file(&patched).unwrap();
}
