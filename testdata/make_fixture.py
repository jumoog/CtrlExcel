"""Generate CtrlExcelReaderFixture.xlsx for ExcelRoundTripTest.ctl.

The extension's own writer cannot produce the cases that only occur in files
from Excel or other tools, so this script writes the workbook XML directly,
following Excel's conventions (shared strings, _xHHHH_ escapes, error cells,
formulas without a cached value, integer date serials, custom number formats,
formatted empty rows, missing rows, rows and cells without the optional r
attribute, hidden rows). Only the Python standard library is used.

    python testdata/make_fixture.py

Sheets "Fixture", "Implicit" and "Hidden", columns Label | Value; the expected read
results are checked by excelFixtureTest() in ExcelRoundTripTest.ctl.
"""

import os
import zipfile
from xml.sax.saxutils import escape

HERE = os.path.dirname(os.path.abspath(__file__))
OUTPUT = os.path.join(HERE, 'CtrlExcelReaderFixture.xlsx')

# Style indices in cellXfs below.
STYLE_DATE = 1      # built-in number format 14 (short date)
STYLE_DURATION = 2  # custom number format 164 "[h]:mm"

shared_strings = []


def sst(text):
    """Index of text in the shared string table (text is stored verbatim)."""
    if text not in shared_strings:
        shared_strings.append(text)
    return shared_strings.index(text)


def label(row):
    return f'<c r="A{row}" t="s"><v>{sst(LABELS[row])}</v></c>'


LABELS = {
    1: 'Label', 2: 'pre1970', 3: 'dateInt', 4: 'duration', 5: 'error',
    6: 'uncached', 7: 'crlf', 8: 'literal', 9: 'exponent', 10: 'bigint',
    12: 'after',
}

# (row, value cell XML) -- the expected CTRL values are in the .ctl test.
VALUE_CELLS = {
    1:  f'<c r="B1" t="s"><v>{sst("Value")}</v></c>',
    # 1960-05-01, outside CTRL's time range -> Excel serial as float
    2:  f'<c r="B2" s="{STYLE_DATE}"><v>22037</v></c>',
    # 2026-01-01 stored without '.', typed Integer by OpenXLSX
    3:  f'<c r="B3" s="{STYLE_DATE}"><v>46023</v></c>',
    # 36:00 as elapsed duration -> number 1.5
    4:  f'<c r="B4" s="{STYLE_DURATION}"><v>1.5</v></c>',
    5:  '<c r="B5" t="e"><f>1/0</f><v>#DIV/0!</v></c>',
    # formula never calculated: no <v>
    6:  '<c r="B6"><f>1+1</f></c>',
    # Excel escapes CR as _x000D_
    7:  f'<c r="B7" t="s"><v>{sst("a_x000D_" + chr(10) + "b")}</v></c>',
    # printable escape as written verbatim by a non-Excel tool
    8:  f'<c r="B8" t="s"><v>{sst("AB_x0041_7")}</v></c>',
    9:  '<c r="B9"><v>1E+20</v></c>',
    10: '<c r="B10"><v>12345678901234567890</v></c>',
    12: f'<c r="B12" t="s"><v>{sst("last")}</v></c>',
}


def sheet_xml():
    rows = []
    for r in range(1, 13):
        if r == 11:
            # formatted but empty row: counted by rowCount(), no values
            rows.append('<row r="11" ht="30" customHeight="1"/>')
            continue
        rows.append(f'<row r="{r}">{label(r)}{VALUE_CELLS[r]}</row>')
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
            '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
            'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
            '<dimension ref="A1:B12"/>'
            f'<sheetData>{"".join(rows)}</sheetData>'
            '</worksheet>')


def implicit_sheet_xml():
    """Row 3 is missing; rows 5 and 6 and some cells omit r (ECMA-376 allows
    it: the element follows the previous one); row 5 ends in <extLst>."""
    def text(t, ref=None):
        r = f' r="{ref}"' if ref else ''
        return f'<c{r} t="s"><v>{sst(t)}</v></c>'
    rows = [
        f'<row r="1">{text("Label", "A1")}{text("Value", "B1")}</row>',
        f'<row r="2">{text("explicit", "A2")}<c r="B2"><v>1</v></c></row>',
        f'<row r="4">{text("afterGap", "A4")}<c r="B4"><v>2</v></c></row>',
        f'<row>{text("implicitRow")}<c><v>3</v></c><extLst/></row>',
        f'<row>{text("implicitCell", "A6")}<c><v>4</v></c></row>',
    ]
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
            '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
            'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
            '<dimension ref="A1:B6"/>'
            f'<sheetData>{"".join(rows)}</sheetData>'
            '</worksheet>')


def hidden_sheet_xml():
    """Hidden rows as Excel writes them: the header row (still the header),
    a data row (3) and a row without cells (4)."""
    def text(t, ref):
        return f'<c r="{ref}" t="s"><v>{sst(t)}</v></c>'
    rows = [
        f'<row r="1" hidden="1">{text("Label", "A1")}{text("Value", "B1")}</row>',
        f'<row r="2">{text("shown", "A2")}<c r="B2"><v>1</v></c></row>',
        f'<row r="3" hidden="1">{text("hidden", "A3")}<c r="B3"><v>2</v></c></row>',
        '<row r="4" hidden="1"/>',
        f'<row r="5">{text("shownAfter", "A5")}<c r="B5"><v>3</v></c></row>',
    ]
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
            '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
            'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
            '<dimension ref="A1:B5"/>'
            f'<sheetData>{"".join(rows)}</sheetData>'
            '</worksheet>')


def shared_strings_xml():
    items = ''.join(f'<si><t xml:space="preserve">{escape(t)}</t></si>' for t in shared_strings)
    n = len(shared_strings)
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
            '<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
            f'count="{n}" uniqueCount="{n}">{items}</sst>')


STYLES = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
<numFmts count="1"><numFmt numFmtId="164" formatCode="[h]:mm"/></numFmts>
<fonts count="1"><font><sz val="11"/><name val="Calibri"/><family val="2"/></font></fonts>
<fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills>
<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>
<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
<cellXfs count="3">
<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>
<xf numFmtId="14" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>
<xf numFmtId="164" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>
</cellXfs>
<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>
</styleSheet>'''

WORKBOOK = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
<sheets><sheet name="Fixture" sheetId="1" r:id="rId1"/><sheet name="Implicit" sheetId="2" r:id="rId4"/><sheet name="Hidden" sheetId="3" r:id="rId5"/></sheets>
</workbook>'''

WORKBOOK_RELS = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/>
<Relationship Id="rId4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet2.xml"/>
<Relationship Id="rId5" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet3.xml"/>
</Relationships>'''

ROOT_RELS = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>
<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>
</Relationships>'''

CONTENT_TYPES = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Default Extension="xml" ContentType="application/xml"/>
<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
<Override PartName="/xl/worksheets/sheet2.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
<Override PartName="/xl/worksheets/sheet3.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
<Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>
<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>
<Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>
</Types>'''

APP = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">
<Application>Microsoft Excel</Application>
<HeadingPairs><vt:vector size="2" baseType="variant"><vt:variant><vt:lpstr>Worksheets</vt:lpstr></vt:variant><vt:variant><vt:i4>3</vt:i4></vt:variant></vt:vector></HeadingPairs>
<TitlesOfParts><vt:vector size="3" baseType="lpstr"><vt:lpstr>Fixture</vt:lpstr><vt:lpstr>Implicit</vt:lpstr><vt:lpstr>Hidden</vt:lpstr></vt:vector></TitlesOfParts>
</Properties>'''

CORE = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
<dc:creator>CtrlExcelReader fixture</dc:creator>
</cp:coreProperties>'''


def main():
    sheet = sheet_xml()  # fills the shared string table
    implicit = implicit_sheet_xml()
    hidden = hidden_sheet_xml()
    parts = {
        '[Content_Types].xml': CONTENT_TYPES,
        '_rels/.rels': ROOT_RELS,
        'docProps/app.xml': APP,
        'docProps/core.xml': CORE,
        'xl/workbook.xml': WORKBOOK,
        'xl/_rels/workbook.xml.rels': WORKBOOK_RELS,
        'xl/styles.xml': STYLES,
        'xl/sharedStrings.xml': shared_strings_xml(),
        'xl/worksheets/sheet1.xml': sheet,
        'xl/worksheets/sheet2.xml': implicit,
        'xl/worksheets/sheet3.xml': hidden,
    }
    with zipfile.ZipFile(OUTPUT, 'w', zipfile.ZIP_STORED) as z:
        for name, data in parts.items():
            # Fixed timestamp and no compression (deflate output depends on the
            # zlib build) keep the output byte-identical on every machine.
            info = zipfile.ZipInfo(name, date_time=(2026, 1, 1, 0, 0, 0))
            info.compress_type = zipfile.ZIP_STORED
            z.writestr(info, data.encode('utf-8'))
    print(f'wrote {OUTPUT}')


if __name__ == '__main__':
    main()
