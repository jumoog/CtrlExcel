# AGENTS.md

Guidance for coding agents working in this repository. User-facing API docs live in `README.md`.

## What this is

A WinCC OA CTRL extension (`CtrlExcelReader`) that reads and writes `.xlsx` files. The file work is done by a Rust library in `rust/` (calamine reads, rust_xlsxwriter writes), linked statically through a small C ABI; the C++ side converts between cells and WinCC OA `Variable`s. It exposes five CTRL functions (`excelGetSheetNames`, `excelReadSheet`, `excelReadFile`, `excelWriteSheet`, `excelWriteFile`), each also as a non-blocking `...Async` variant that runs on a worker pool and delivers through a reference parameter. Read the threading rules under *CTRL extension conventions* before touching `ExcelExternHdl.cxx`.

| File | Role |
|---|---|
| `ExcelExternHdl.cxx/.hxx` | CTRL entry point: `fnList` signatures, argument evaluation, error reporting; the operations (`getSheetNames`, `readSheet`, `readFile`, `writeWorkbook`), `PathLock`, `WorkerPool` and `JobWait` for the `...Async` variants |
| `ExcelXlsxHelpers.cxx/.hxx` | Cell ⇄ WinCC OA `Variable` conversion, header keys, encoding conversion, `readSheetRows` (`RowBuilder`) / `writeSheetData` |
| `ExcelXlsxCore.cxx/.hxx` | Everything that needs the Rust library but not the WinCC OA API: calendar/date conversion, number typing (`integerFit`), header texts, case folding, sheet-name checks, and the RAII wrappers `Workbook` (+ `SheetVisitor`) and `WorkbookWriter` over the C ABI |
| `rust/include/ctrlexcel_xlsx.h`, `rust/src/ffi.rs` | The C ABI (keep both in sync) |
| `rust/src/reader.rs`, `scan.rs`, `writer.rs` | Reading (calamine's streaming cell reader merged with our own row scan), writing (rust_xlsxwriter, temp file + rename) |
| `rust/tests/read_write.rs` | Rust tests (fixture, round trips, hidden rows, 1904 dates; run in CI) |
| `tests/ExcelXlsxCoreTest.cxx` | C++ tests of `ExcelXlsxCore`, including the library through the C ABI (no WinCC OA needed; run in CI) |
| `ExcelRoundTripTest.ctl` | CTRL round-trip tests (need a WinCC OA project to run) |
| `scripts/libs/classes/ExcelWorkbook.ctl` | CTRL++ classes `ExcelWorkbook` (file, load/save, messages, call mode) and `ExcelSheet` (rows in memory) on top of the functions |
| `ExcelWorkbookTest.ctl` | CTRL tests of those classes |
| `docker/` | Image on top of the WinCC OA image that builds the extension and runs every test suite on Linux (`run-tests.sh`) |
| `CMakeLists.txt`, `cmake/RustLibrary.cmake`, `CMakePresets.json`, `.github/workflows/ci.yml` | Build (CMake runs cargo), CI; the version is in `CMakeLists.txt` (`project(... VERSION)`) only |
| `VersInfo.rc`, `cmake/VersionInfo.cmake` | Windows version resource; the script regenerates `CtrlExcelVersion.h` on every build with the git revision, so `ProductVersion` is e.g. `1.0.0+a5f43da` (`-dirty` and `PrivateBuild` for uncommitted changes) |
| `rust/third_party_licenses.py` | Regenerates `THIRD-PARTY-LICENSES.txt` from `rust/Cargo.lock`; run it after dependency changes |
| `.github/dependabot.yml` | Weekly updates of the crates (`rust/Cargo.lock`, one grouped PR), the toolchain (`rust/rust-toolchain.toml`, which docker/ follows) and the GitHub Actions. Before merging a crate update, regenerate `THIRD-PARTY-LICENSES.txt` and run the CTRL tests (CI cannot) |

## Build

Requires `API_ROOT` (WinCC OA `api/` dir) and Rust via rustup (the toolchain is pinned in `rust/rust-toolchain.toml`; cargo must run in `rust/` for rustup to pick it up, which `cmake/RustLibrary.cmake` does).

```powershell
cmake --preset windows
cmake --build --preset relwithdebinfo
```

Output: `build/RelWithDebInfo/CtrlExcelReader.dll`. Linux presets: `linux-relwithdebinfo` / `linux-debug`. C++17 is set by the WinCC OA `CMakeDefines.txt`. `-DCTRLEXCEL_BUILD_EXTENSION=OFF` builds only the C++ tests, without `API_ROOT` (this is what CI does); `-DCTRLEXCEL_BUILD_TESTS=OFF` skips them. cargo always builds the release profile into `<build>/cargo`.

Always build after changing C++ or Rust code and keep it warning-free (CI builds Rust with `RUSTFLAGS=-D warnings`).

## Testing

**Rust tests** (`cd rust; cargo test --locked`) cover the library: reading the fixture, layout (rows, hidden rows, extent), escapes, numbers, dates, the 1904 system, writing and replacing files. Add a test there for new file-level behaviour.

**C++ tests** (`tests/ExcelXlsxCoreTest.cxx`, plain `CHECK`/`CHECK_EQ` macros, no framework) cover `ExcelXlsxCore` and the library through the C ABI. Run them after every build: `ctest --test-dir build -C RelWithDebInfo --output-on-failure`. GitHub Actions builds both on Linux and Windows with warnings as errors. Keep logic that does not need WinCC OA in `ExcelXlsxCore` (never include WinCC OA headers there, or CI cannot build it). Tests set the time zone themselves (`setTimeZone`); DST cases run on POSIX only, since MSVC applies US DST rules to any `TZ` value. The tests read the fixture (path compiled in via `tests/FixturePath.hxx.in`; an argument overrides it), so new fixture rows can be checked in `fixtureTest()` as well. CI cannot build `ExcelXlsxHelpers.cxx` or `ExcelExternHdl.cxx`, so a green CI run does not replace the local extension build.

**CTRL tests**: `ExcelRoundTripTest.ctl` runs inside a WinCC OA project with the built DLL in `bin/` (the script loads it with `#uses "CtrlExcelReader"`; `LoadCtrlLibs` is for CTRL script libraries and does not load extensions); each test logs `pass` via `DebugTN`. Without a running project, `WCCOActrl -proj <project> -n -log +stderr ExcelRoundTripTest.ctl` runs it without Data/Event. On Linux, `docker/` builds the extension and runs all suites (see README, *Build and test on Linux in Docker*); without Docker or a WinCC OA project, say so and ask the user to run it after behaviour changes. After every rebuild, the new DLL must be copied into the project's `bin/` and the CTRL manager restarted (a running manager keeps the old DLL loaded). A test failure that contradicts the current code is most likely a stale DLL: before debugging, compare the project DLL's `ProductVersion` (`(Get-Item bin/CtrlExcelReader.dll).VersionInfo.ProductVersion`, commit after the `+`) with `git describe --always --dirty`. Add a test function there (and call it from `main()`) for new behaviour. Cases the extension's writer cannot produce (files from Excel or other tools) go into the fixture: extend `testdata/make_fixture.py`, regenerate `testdata/CtrlExcelReaderFixture.xlsx` (`python testdata/make_fixture.py`, byte-reproducible), and check the new row in `excelFixtureTest()` and the Rust tests. The user must copy the fixture into the project's `data/` directory; the test skips if it is missing. Register every test in `main()` via `recordTest()`, and call `skipTest()` before returning TRUE from a test that cannot run, so the final `ExcelRoundTripTest summary` line counts it as skipped rather than passed.

CTRL `time` cannot go below 1970-01-01 UTC (`makeTime` returns time 0, arithmetic clamps at 0), and its maximum is `TimeVar::MaxTimeVarSec` (2262 on 64-bit). Tests using `makeTime` for dates outside that range test nothing. In tests, store `getLastError()` in a variable right after the call under test: any other function call (even `dynlen`) replaces the error list.

## Known pitfalls

- **Linux symbol clashes**: the CTRL manager loads extensions into one global symbol namespace, and `libCtrlV321` exports a class `ExternHdl` (hence `ExcelExternHdl`; a class of that name in the extension got WinCC OA's vtable, so none of its functions existed). The extension is built with hidden visibility and `--exclude-libs,ALL` (`CMakeLists.txt`), which also hides every symbol of the Rust static library, so it exports only `newExternHdl` / `getVersionString`; keep it that way and do not name classes or headers after WinCC OA's (`ExternHdl.hxx` also shares the API's include guard).
- **C ABI**: no panic or C++ exception may cross it. Every `extern "C"` function in `ffi.rs` runs under `catch_unwind` (`guarded`) and reports errors through `cxlsx_error`; the C++ visitor callbacks (`VisitorThunk`) catch everything, keep it as `exception_ptr` and rethrow after the library returns. Allocation failure in Rust aborts the process (not catchable), so never allocate in proportion to the sheet's extent.
- **Never call calamine's `worksheet_range()`** (or anything built on `Range`) on files we did not write: it allocates rows × columns cells, so a stray cell at XFD1048576 asks for 550 GB and aborts the manager. Reading goes through `worksheet_cells_reader()` (`next_cell_with_formula_metadata`), which streams.
- **What calamine does not expose** is read by `scan.rs` directly from the package: the sheet's part path, the 1904 date system (`workbookPr date1904`; the reader adds 1462 so C++ always gets 1900-based serials), and per sheet the stored rows, their `hidden` flag and the highest column (also of cells without a value). Rows and cells without the optional `r` attribute follow the previous one; ones that do not ascend are skipped. `reader.rs` merges these rows with calamine's cells so the visitor sees `row()` before the cells of that row, also for rows without cells. The scan runs in addition to calamine's pass, so it is a byte scanner (memchr from tag to tag, chunked, aware of quotes, comments and CDATA) rather than an XML parser: quick-xml took longer than calamine's whole cell parse.
- **Windows C runtime**: Rust always links the release CRT (`msvcrt`). WinCC OA's `CMakeDefines.txt` uses `/MD` in Debug builds too, and `CMakeLists.txt` sets `CMAKE_MSVC_RUNTIME_LIBRARY` to match for the tests-only build; never link `/MDd` code with the library. The system libraries the static library needs come from `cargo rustc -- --print native-static-libs`; update the list there when dependencies change.
- **Writing**: `WorkbookWriter::save` writes a temporary file next to the target and renames it over it, so a failed write never clobbers an existing file. The rename must keep what writing in place kept: it follows symbolic links (also to a file that does not exist yet) and gives the new file the old one's permissions on Unix (`writer.rs`, Unix-only Rust tests; owner and ACLs are not kept); validate everything before writing cells, and skip `save()` on failure. Sheets use rust_xlsxwriter's low-memory mode: rows go to a temporary file as they are written, so they must be written in ascending order (`writeSheetData` does). Columns are the union of all rows' keys; look up values with the original key `Variable`s, not stringified `TextVar`s, or non-text keys (e.g. int column keys) miss. Texts go through `toCellText()` (valid UTF-8 via `toValidUtf8()`, 32767-character limit counted in UTF-16 units); rust_xlsxwriter itself escapes XML-invalid characters, CR and literal `_xHHHH_` text as `_xHHHH_`. Empty texts are not stored (Excel does not either). Non-finite numbers become the error `#NUM!`. Sheet names go through `checkSheetNames()` (Unicode-aware case folding via `foldCase()`, `History`). Column widths are capped at 255.
- **Escapes on read**: calamine decodes every `_xHHHH_` escape (as ECMA-376 specifies), also printable ones other tools write verbatim (`AB_x0041_7` reads as `ABA7`); U+0000 is dropped in `reader.rs` because CTRL strings are C strings.
- **Dates**: `serialToLocalTime()` / `localTimeToSerial()` use pure calendar arithmetic (`daysFromCivil`, Excel's fictitious 1900-02-29 handled in `excelDayToUnixDay`), and use `mktime`/`localtime` only for the UTC offset, which `ReadState::dayOffset` caches per calendar day (MSVC's mktime is slow); only DST transition days are converted per cell, and nonexistent times there are counted as shifted. calamine decides from the number format whether a number is a date (`CXLSX_DATETIME`) or a duration (`CXLSX_DURATION`: `[h]:mm` and the like stay numbers); `isDateCell()` additionally keeps serials below 1 (time-only values) as numbers. Reads return the serial as float (with a warning) for dates outside CTRL's time range; writes need no range check, since every CTRL time is a valid Excel date. The DST fall-back hour is inherently ambiguous (documented in README).
- **Numbers**: Excel stores doubles, and calamine returns every number as `f64`. `integerFit()` makes whole numbers `int`, or `LongVar` beyond 32 bits (never a truncated `IntegerVar`), everything else `float`; integers beyond 2^53 cannot round-trip exactly.
- **Building rows** (`RowBuilder`): use the capturing overloads `MappingVar::setAt(Variable *, Variable *)` (via `put()`) and `DynVar::append(VariablePtr)`; the `const Variable &` ones clone every value and every finished row once more (about 0.07 s per 100,000 rows × 10 columns). The blocking functions' result is copied by CTRL anyway (about 0.2 s for that sheet; the API returns `const Variable *`), the `...Async` variants move it.
- **Header keys** (`RowBuilder`): text for non-empty headers, repeated names suffixed `_2`, `_3`…, empty headers fall back to the 1-based column number as `int`. Numeric, boolean and date headers go through `headerText()`.
- **Encoding**: CTRL strings are in the project encoding (UTF-8 or e.g. ISO-8859-1); the library and the file use UTF-8. Convert every filename, sheet name and text with `ExcelXlsxHelpers::toUtf8()` / `fromUtf8()` (WinCC OA's `UTF8Converter`). The C ABI takes UTF-8 paths, which Rust converts to the OS's path encoding itself. Build `std::filesystem` paths from the UTF-8 name with `std::filesystem::u8path`.
- **Reading rows**: rows without any value are skipped by default (`skipEmptyRows`), because a sheet's rows include rows that only carry formatting. A formula without a cached value arrives as `CXLSX_UNCACHED_FORMULA` and reads as `""` with a warning. Non-fatal problems go into readSheetRows' `warnings` and are reported by the caller.

## CTRL extension conventions

- `ExecuteParamRec` is nested in `BaseExternHdl`; helpers outside the class take `CtrlThread *`.
- Every function starts with `param.thread->clearLastError()`, then `hasNumArgs(min, max, param)`.
- Evaluate arguments with `evalArg` / `evalOptionalArg`: `evaluate()` may return `nullptr`.
- Never swallow exceptions silently: collect them in `Messages` (`add` / `addCurrentException`) and report them with `Messages::report` / `reportError`, so scripts see them via `getLastError()`. Include the file (and sheet) name in the context. Messages are UTF-8 (`utf8Arg()` for CTRL arguments); `reportError` converts them to the project encoding, so never mix in project-encoded text.
- Results are returned via function-local `static` variables, reset at the start of each call.
- **Blocking vs. `...Async`**: all work lives in the operations in `ExcelExternHdl.cxx` (`getSheetNames`, `readSheet`, `readFile`, `writeWorkbook`), which take UTF-8 inputs and collect errors in `Messages` instead of touching the `CtrlThread`. The blocking functions call them inline; the `...Async` variants queue them on the `WorkerPool` (2 to 4 threads, joined at exit) via `startJob`, and a `JobWait` (`WaitCond`) delivers the result into the reference parameter and reports the messages in `checkDone()`, on the CTRL thread. Callers hold the file's `PathLock` (per lexically normalised path; shared for reads, exclusive for writes) because async jobs really run in parallel; it only try-locks: the pool picks the first queued job whose file is free, and blocking calls fail with a "busy" error instead of freezing the manager behind an async job. Operations produce UTF-8 only; the `finish*` / `deliverSheets` steps and `convertRowsToProjectEncoding` convert to the project encoding on the CTRL thread, since WinCC OA's `UTF8Converter` is not known to be thread-safe. Rules: never call `reportError`, `appendLastError`, `resolveTarget`/`getTarget` or `toUtf8`/`fromUtf8`/`isRepresentable` from a worker; clone write inputs before queueing; queue first and only then `setWaitCond` (a failed submit must not leave the script waiting); resolve the target again at delivery instead of keeping the pointer; hand rows over with `moveAllItems` (not a deep copy on the CTRL thread); nothing may throw out of `checkDone`; keep job state in the shared `Job`, since the script (and its `JobWait`) may go away while the worker runs. WinCC OA's `Allocator` falls back to plain `new`/`delete` off the main thread, so building `Variable`s in the worker is safe. A `Workbook` or `WorkbookWriter` is used by one thread at a time; the Rust library has no global state.
- Keep `fnList` signatures in sync with the README.

## CTRL++ (`ExcelWorkbook.ctl`)

- WinCC OA 3.21 rejects `const` member functions (the error is reported at the using script's `main`).
- `vec[i] = ptr` on a `vector<shared_ptr<T>>` copies into the object `vec[i]` points to (every holder of that pointer sees the change); replace an element with `removeAt` + `insertAt`.
- Mapping key order is unspecified (neither insertion nor sorted order), so never rely on it; `load()` takes the sheet order from `excelGetSheetNames`.
- Store `getLastError()` in the class right after each extension call; capture `...Async` results in local variables, not members, since the reference argument is evaluated again on delivery.
- The extension reports errors and warnings alike as `PRIO_WARNING`; the classes decide success by the result (empty read with messages = failure).

## Style

- Match the surrounding file: `ExcelExternHdl.cxx` uses `if ( cond )` with inner spaces; `ExcelXlsxHelpers.cxx` uses `if (cond)`. Two-space indent, braces on their own line. Rust: `cargo fmt` style.
- Comments explain *why* (library quirks, invariants), not what.
- Update `README.md` when user-visible behaviour changes.

## Git

- Conventional commits: `type(scope): summary`, e.g. `fix(excel): ...`, `refactor(excel): ...`, `test(excel): ...`, `docs: ...`, `fix(build): ...`.
- `build/` and `rust/target/` are ignored; never commit build output. Commit `rust/Cargo.lock`.
