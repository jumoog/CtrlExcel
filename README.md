# WinCC OA CTRL Extension for Excel

A WinCC OA CTRL extension that adds `.xlsx` file reading and writing capabilities using [OpenXLSX](https://github.com/troldal/OpenXLSX).

## CTRL Functions

### `excelGetSheetNames`

```ctrl
dyn_string excelGetSheetNames(string filename)
```

Returns the names of all worksheets in the given `.xlsx` file. Chart sheets are not listed, since they contain no cells to read.

```ctrl
dyn_string sheets = excelGetSheetNames("C:/data/report.xlsx");
// sheets = {"Sheet1", "Sheet2", "Summary"}
```

### `excelReadSheet`

```ctrl
dyn_mapping excelReadSheet(string filename, string sheetName, bool skipHiddenRows = TRUE, bool firstRowIsColumnNames = TRUE)
```

Reads a sheet and returns each data row as a `mapping`.

- **sheetName** — pass an empty string to read the first sheet.
- **skipHiddenRows** — optional, defaults to `TRUE`. When `TRUE`, hidden rows are omitted.
- **firstRowIsColumnNames** — optional, defaults to `TRUE`. When `TRUE`, the first row supplies the mapping keys as strings. When `FALSE`, keys are 1-based column integers.

Header cells become string keys: numbers are written without decimals (`2024`), dates as `YYYY-MM-DD` (plus ` HH:MM:SS` when they have a time part). An empty header cell falls back to its 1-based column number (`int` key). A repeated header gets a suffix (`Value`, `Value_2`, `Value_3`) so no column is overwritten.

Cell values are automatically typed based on the Excel cell type:

- Numeric integers → `int`, or `long` outside the 32-bit range
- Numeric decimals → `float`. Excel does not distinguish integers from decimals, so a whole-number `float` such as `87.0` is stored as `87` and reads back as `int`.
- Booleans → `bool`
- Dates/times → `time` with milliseconds (Excel serial converted to WinCC OA time, treated as local time). Time-only values (serial below 1, e.g. `08:00`) are returned as numbers.
- Strings → `string`
- Error cells (e.g. a formula result `#DIV/0!` or `#N/A`) → the error code as `string`, so they are distinguishable from empty cells
- Empty cells → `""`

Excel date serials carry no timezone. In the hour repeated when daylight saving time ends, a written time cannot be told apart from the same wall-clock time one hour earlier and may read back one hour off. Times the C runtime cannot convert (before 1970 on Windows, after the year 3000) use the local standard-time offset without daylight saving time, in both directions, so they still round-trip.

```ctrl
// Read the first sheet with default options
dyn_mapping rows = excelReadSheet("C:/data/report.xlsx", "");

// Read a specific sheet, include hidden rows
dyn_mapping rows = excelReadSheet("C:/data/report.xlsx", "Sheet2", FALSE);

// Given an Excel sheet:
//   | Name  | Age | Score |
//   | Alice | 30  | 95.5  |
//   | Bob   | 25  | 87.0  |

// With firstRowIsColumnNames = TRUE (default):
DebugN(rows[1]["Name"]);   // "Alice"  (string)
DebugN(rows[1]["Age"]);    // 30       (int)
DebugN(rows[1]["Score"]);  // 95.5     (float)

// With firstRowIsColumnNames = FALSE:
dyn_mapping rows = excelReadSheet("C:/data/report.xlsx", "", TRUE, FALSE);
DebugN(rows[1][1]);  // "Name"   (string)
DebugN(rows[1][2]);  // "Age"    (string)
DebugN(rows[2][1]);  // "Alice"  (string)
DebugN(rows[2][2]);  // 30       (int)
```

### `excelReadFile`

```ctrl
mapping excelReadFile(string filename, bool skipHiddenRows = TRUE, bool firstRowIsColumnNames = TRUE)
```

Reads all sheets in the file at once. Returns a `mapping` where keys are sheet names and values are `dyn_mapping` (rows).

- **skipHiddenRows** — optional, defaults to `TRUE`.
- **firstRowIsColumnNames** — optional, defaults to `TRUE`.

```ctrl
mapping allSheets = excelReadFile("C:/data/report.xlsx");

// Access sheets by name
DebugN(allSheets["Sheet1"][1]["Name"]);   // first row of Sheet1
DebugN(allSheets["Summary"][1]["Total"]); // first row of Summary sheet
```

### `excelWriteSheet`

```ctrl
bool excelWriteSheet(string filename, string sheetName, dyn_anytype data)
```

Writes a `dyn_anytype` (containing mappings) to a single-sheet `.xlsx` file. The keys of all rows, in order of first appearance, become the column headers; a row without a key gets an empty cell in that column.

- An empty **sheetName** writes `Sheet1`. Sheet names must follow Excel's rules: at most 31 characters, none of `\ / ? * [ ] :`, no leading or trailing apostrophe, not `History`, and unique ignoring case.
- Control characters other than tab, newline and carriage return are removed from texts (XML cannot store them). Bytes that are not valid UTF-8 are treated as ISO-8859-1 and converted, so projects running a Latin-1 codepage still produce valid files. A text longer than 32767 characters, Excel's cell limit, fails the write.
- More than 1048575 data rows or 16384 columns (Excel's sheet size) fail the write. Column widths follow the longest value, capped at Excel's maximum of 255.

```ctrl
dyn_mapping rows;
mapping row1, row2;
row1["Name"] = "Alice";
row1["Age"]  = 30;
row2["Name"] = "Bob";
row2["Age"]  = 25;
dynAppend(rows, row1);
dynAppend(rows, row2);

bool ok = excelWriteSheet("C:/data/output.xlsx", "People", rows);
```

### `excelWriteFile`

```ctrl
bool excelWriteFile(string filename, mapping data)
```

Writes a `mapping` where keys are sheet names and values are `dyn_anytype` (rows of mappings) to a multi-sheet `.xlsx` file. This is the same format returned by `excelReadFile`.

```ctrl
// Read and write back (round-trip)
mapping allSheets = excelReadFile("C:/data/input.xlsx");
bool ok = excelWriteFile("C:/data/output.xlsx", allSheets);

// Build from scratch
dyn_mapping rows;
mapping row;
row["Name"] = "Alice";
dynAppend(rows, row);

mapping data;
data["Sheet1"] = rows;
bool ok = excelWriteFile("C:/data/output.xlsx", data);
```

## Error handling

Every function clears and then fills the CTRL error list, so failures can be inspected with `getLastError()`:

- Read functions return an empty result on failure (missing file, unknown sheet, corrupt workbook). `excelReadFile` still returns the sheets it could read and reports the ones it could not.
- Write functions return `FALSE` on failure (file open in Excel, invalid sheet name, data that is not a `dyn_mapping` of mappings, text over the cell limit, data larger than a sheet). A failed write never modifies an existing file.

```ctrl
dyn_mapping rows = excelReadSheet("C:/data/report.xlsx", "Sheet1");
dyn_errClass err = getLastError();
if (dynlen(err) > 0)
  DebugTN("excelReadSheet failed", err);
```

## Build

### Prerequisites

- WinCC OA 3.20 API (set `API_ROOT` to your WinCC OA `api/` directory)
- [vcpkg](https://github.com/microsoft/vcpkg) (set `VCPKG_ROOT`)
- CMake 3.15+
- Windows: Visual Studio (uses the `vs2022-vcpkg` preset)
- Linux: a C++ toolchain + `make` (uses `Unix Makefiles`)

This repository uses **vcpkg manifest mode** (`vcpkg.json`). Dependencies are
resolved automatically when you configure with the vcpkg toolchain.

### Configure and build (Windows)

PowerShell:

```powershell
$env:API_ROOT = "C:\Siemens\Automation\WinCC_OA\3.20\api"

cmake --preset vs2022-vcpkg
cmake --build --preset relwithdebinfo
```

Debug build:

```powershell
cmake --build --preset debug
```

### Configure and build (Linux)

```sh
export API_ROOT=/opt/WinCC_OA/3.20/api

cmake --preset linux-vcpkg-relwithdebinfo
cmake --build --preset linux-relwithdebinfo
```

Debug build:

```sh
cmake --preset linux-vcpkg-debug
cmake --build --preset linux-debug
```

### Install

Copy the built library into your WinCC OA project's `bin/` directory:

- Windows: `build/RelWithDebInfo/CtrlExcelReader.dll` (or `build/Debug/...`)
- Linux: `build-linux-relwithdebinfo/CtrlExcelReader.so` (or `build-linux-debug/...`)

Add the extension to your WinCC OA config file:

```text
[ctrl]
LoadCtrlLibs = "CtrlExcelReader"
```

## Licenses

This extension is licensed under the MIT License (see `LICENSE` file). It uses [OpenXLSX](https://github.com/troldal/OpenXLSX) which is licensed under the BSD 3-Clause License (see `THIRD-PARTY-LICENSES.txt`).
