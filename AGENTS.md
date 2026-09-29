# AGENTS.md

Guidance for coding agents working in this repository. User-facing API docs live in `README.md`.

## What this is

A WinCC OA CTRL extension (`CtrlExcelReader`) that reads and writes `.xlsx` files via OpenXLSX (vcpkg). It exposes five CTRL functions (`excelGetSheetNames`, `excelReadSheet`, `excelReadFile`, `excelWriteSheet`, `excelWriteFile`), each also as a non-blocking `...Async` variant that runs on a worker pool and delivers through a reference parameter. Read the threading rules under *CTRL extension conventions* before touching `ExternHdl.cxx`.

| File | Role |
|---|---|
| `ExternHdl.cxx/.hxx` | CTRL entry point: `fnList` signatures, argument evaluation, error reporting; the operations (`getSheetNames`, `readSheet`, `readFile`, `writeWorkbook`), `PathLock`, `WorkerPool` and `JobWait` for the `...Async` variants |
| `ExcelXlsxHelpers.cxx/.hxx` | Cell ⇄ WinCC OA `Variable` conversion, header keys, encoding conversion, `readSheetRows` / `writeSheetData` |
| `ExcelXlsxCore.cxx/.hxx` | Everything that needs OpenXLSX but not the WinCC OA API: calendar/date conversion, `_xHHHH_` escaping, case folding, sheet-name checks, cell formats, date-format detection, `cellNumber`, `headerText`, `forEachCell` |
| `tests/ExcelXlsxCoreTest.cxx` | C++ tests of `ExcelXlsxCore` (no WinCC OA needed; run in CI) |
| `ExcelRoundTripTest.ctl` | CTRL round-trip tests (need a WinCC OA project to run) |
| `CMakeLists.txt`, `CMakePresets.json`, `vcpkg.json`, `.github/workflows/ci.yml` | Build, CI |

## Build

Requires `API_ROOT` (WinCC OA `api/` dir) and `VCPKG_ROOT` in the environment.

```powershell
cmake --preset vs2022-vcpkg
cmake --build --preset relwithdebinfo
```

Output: `build/RelWithDebInfo/CtrlExcelReader.dll`. Linux presets: `linux-vcpkg-relwithdebinfo` / `linux-relwithdebinfo`. C++17 is set by the WinCC OA `CMakeDefines.txt`. `-DCTRLEXCEL_BUILD_EXTENSION=OFF` builds only the C++ tests, without `API_ROOT` (this is what CI does); `-DCTRLEXCEL_BUILD_TESTS=OFF` skips them.

Always build after changing C++ code and keep it warning-free.

## Testing

**C++ tests** (`tests/ExcelXlsxCoreTest.cxx`, plain `CHECK`/`CHECK_EQ` macros, no framework) cover `ExcelXlsxCore` and the OpenXLSX quirks it works around. Run them after every build: `ctest --test-dir build -C RelWithDebInfo --output-on-failure`. GitHub Actions builds them on Linux and Windows with warnings as errors. Add a test there for new logic that does not need WinCC OA, and keep such logic in `ExcelXlsxCore` (never include WinCC OA headers there, or CI cannot build it). Tests set the time zone themselves (`setTimeZone`); DST cases run on POSIX only, since MSVC applies US DST rules to any `TZ` value. The tests read the fixture (path compiled in via `tests/FixturePath.hxx.in`; an argument overrides it), so new fixture rows can be checked in `fixtureTest()` as well. CI cannot build `ExcelXlsxHelpers.cxx` or `ExternHdl.cxx`, so a green CI run does not replace the local extension build.

**CTRL tests**: `ExcelRoundTripTest.ctl` runs inside a WinCC OA project with the built DLL loaded (`LoadCtrlLibs = "CtrlExcelReader"`); each test logs `pass` via `DebugTN`. An agent cannot run it: say so and ask the user to run it after behaviour changes. After every rebuild, the new DLL must be copied into the project's `bin/` and the CTRL manager restarted (a running manager keeps the old DLL loaded). A test failure that contradicts the current code is most likely a stale DLL: ask for the DLL's timestamp before debugging. Add a test function there (and call it from `main()`) for new behaviour. Cases the extension's writer cannot produce (files from Excel or other tools) go into the fixture: extend `testdata/make_fixture.py`, regenerate `testdata/CtrlExcelReaderFixture.xlsx` (`python testdata/make_fixture.py`, byte-reproducible), and check the new row in `excelFixtureTest()`. The user must copy the fixture into the project's `data/` directory; the test skips if it is missing. Register every test in `main()` via `recordTest()`, and call `skipTest()` before returning TRUE from a test that cannot run, so the final `ExcelRoundTripTest summary` line counts it as skipped rather than passed.

CTRL `time` cannot go below 1970-01-01 UTC (`makeTime` returns time 0, arithmetic clamps at 0), and its maximum is `TimeVar::MaxTimeVarSec` (2262 on 64-bit). Tests using `makeTime` for dates outside that range test nothing. In tests, store `getLastError()` in a variable right after the call under test: any other function call (even `dynlen`) replaces the error list.

## Known pitfalls

- **`ssize_t` clash**: WinCC OA's `winnt/win32.h` typedefs `ssize_t` as `int`; OpenXLSX redefines it as 64-bit. Include OpenXLSX only through `ExcelXlsxCore.hxx` (or `ExcelXlsxHelpers.hxx`, which includes it), which renames it via `#define ssize_t OpenXLSX_ssize_t` (Windows only; on Linux there is no clash and the macro could rename glibc's typedef). Never `#include <OpenXLSX.hpp>` directly.
- **OpenXLSX CMake config** does not `find_dependency` its link deps; `CMakeLists.txt` must `find_package` pugixml, miniz and (Windows only, where the vcpkg port links it) Boost nowide before OpenXLSX.
- **`XLDocument::create()`** builds in a temp archive; the target file is only written by `save()`. `writeWorkbook` relies on this: validate everything first, and skip `save()` on failure so an existing file is never clobbered.
- **Dates**: never convert via `XLDateTime::tm()` (truncates to whole seconds) or `XLDateTime(std::tm)`. `serialToLocalTime()` / `localTimeToSerial()` use pure calendar arithmetic (`daysFromCivil`, Excel's fictitious 1900-02-29 handled in `excelDayToUnixDay`), and use `mktime`/`localtime` only for the UTC offset, which `ReadState::dayOffset` caches per calendar day (MSVC's mktime is slow); only DST transition days are converted per cell, and nonexistent times there are counted as shifted. Reads return the serial as float (with a warning) for dates outside CTRL's time range; writes need no range check, since every CTRL time is a valid Excel date. `XLDateTime(serial)` throws for serial < 1, so `isDateNumber()` excludes those. Elapsed-time formats (`[h]`, `[m]`, `[s]`, built-in 46) are durations and stay numbers. The DST fall-back hour is inherently ambiguous (documented in README).
- **Header keys** (`readSheetRows`): text for non-empty headers, repeated names suffixed `_2`, `_3`…, empty headers fall back to the 1-based column number as `int`. Numeric and date headers go through `headerText()`; never `get<std::string>()` on a non-string cell (it throws).
- **Writing**: columns are the union of all rows' keys. Look up values with the original key `Variable`s, not stringified `TextVar`s, or non-text keys (e.g. int column keys) miss. Texts go through `toCellText()` / `toXmlUtf8()` (valid UTF-8, XML-invalid characters and CR as `_xHHHH_` escapes, which `decodeExcelEscapes()` reverses on read; enforces the 32767-character limit); sheet names through `checkSheetNames()` (Unicode-aware case folding via `foldCase()`), because OpenXLSX only rejects exact duplicates. Sheet size (rows/columns) is validated before writing; column widths are capped at 255.
- **Numbers**: OpenXLSX writes doubles with `%.17g` and types any numeric text without `.` as Integer, so whole-number floats read back as `int` (inherent, not a bug to fix). Two consequences are handled in `cellNumber()`: date detection must also run for Integer cells (date-only serials such as `46023`), and exponent text such as `1e+20` is parsed by `as_llong` as `1`, so values -9..9 are re-parsed from the raw XML (`rawNumberText()` via `XLCell::print`). Always go through `cellNumber()`, never `val.get<int64_t>()` directly.
- **Integers**: values outside the 32-bit range must become `LongVar`, not a truncated `IntegerVar`.
- **Encoding**: CTRL strings are in the project encoding (UTF-8 or e.g. ISO-8859-1); OpenXLSX and the file use UTF-8. Convert every filename, sheet name and text with `ExcelXlsxHelpers::toUtf8()` / `fromUtf8()` (WinCC OA's `UTF8Converter`). Build `std::filesystem` paths from the UTF-8 name with `std::filesystem::u8path`.
- **Sparse sheets**: dereferencing OpenXLSX row/cell iterators (`rows()`, `cells()`, `XLCellIterator::operator*`) creates missing rows and cells in the XML. Never iterate a sheet that way; use `rowExists()` on the row iterator and `forEachCell()` (checks `cellExists()` first). A stray cell at row 1048576 or column XFD otherwise materialises millions of nodes.
- **Reading rows**: rows without any value are skipped by default (`skipEmptyRows`), because `rowCount()` includes rows that only carry formatting. Uncached formulas are typed Integer 0 by OpenXLSX; `isUncachedFormula()` detects a missing/empty `<v>`. Non-fatal problems go into readSheetRows' `warnings` and are reported by the caller.

## CTRL extension conventions

- `ExecuteParamRec` is nested in `BaseExternHdl`; helpers outside the class take `CtrlThread *`.
- Every function starts with `param.thread->clearLastError()`, then `hasNumArgs(min, max, param)`.
- Evaluate arguments with `evalArg` / `evalOptionalArg`: `evaluate()` may return `nullptr`.
- Never swallow exceptions silently: collect them in `Messages` (`add` / `addCurrentException`) and report them with `Messages::report` / `reportError`, so scripts see them via `getLastError()`. Include the file (and sheet) name in the context. Messages are UTF-8 (`utf8Arg()` for CTRL arguments); `reportError` converts them to the project encoding, so never mix in project-encoded text.
- Results are returned via function-local `static` variables, reset at the start of each call.
- **Blocking vs. `...Async`**: all work lives in the operations in `ExternHdl.cxx` (`getSheetNames`, `readSheet`, `readFile`, `writeWorkbook`), which take UTF-8 inputs and collect errors in `Messages` instead of touching the `CtrlThread`. The blocking functions call them inline; the `...Async` variants queue them on the `WorkerPool` (2 to 4 threads, joined at exit) via `startJob`, and a `JobWait` (`WaitCond`) delivers the result into the reference parameter and reports the messages in `checkDone()`, on the CTRL thread. Callers hold the file's `PathLock` (per lexically normalised path; shared for reads, exclusive for writes) because async jobs really run in parallel; it only try-locks: the pool picks the first queued job whose file is free, and blocking calls fail with a "busy" error instead of freezing the manager behind an async job. Operations produce UTF-8 only; the `finish*` / `deliverSheets` steps and `convertRowsToProjectEncoding` convert to the project encoding on the CTRL thread, since WinCC OA's `UTF8Converter` is not known to be thread-safe. Rules: never call `reportError`, `appendLastError`, `resolveTarget`/`getTarget` or `toUtf8`/`fromUtf8`/`isRepresentable` from a worker; clone write inputs before queueing; queue first and only then `setWaitCond` (a failed submit must not leave the script waiting); resolve the target again at delivery instead of keeping the pointer; hand rows over with `moveAllItems` (not a deep copy on the CTRL thread); nothing may throw out of `checkDone`; keep job state in the shared `Job`, since the script (and its `JobWait`) may go away while the worker runs. WinCC OA's `Allocator` falls back to plain `new`/`delete` off the main thread, so building `Variable`s in the worker is safe.
- Keep `fnList` signatures in sync with the README.

## Style

- Match the surrounding file: `ExternHdl.cxx` uses `if ( cond )` with inner spaces; `ExcelXlsxHelpers.cxx` uses `if (cond)`. Two-space indent, braces on their own line.
- Comments explain *why* (library quirks, invariants), not what.
- Update `README.md` when user-visible behaviour changes.

## Git

- Conventional commits: `type(scope): summary`, e.g. `fix(excel): ...`, `refactor(excel): ...`, `test(excel): ...`, `docs: ...`, `fix(build): ...`.
- `build/` and `vcpkg_installed/` are ignored; never commit build output.
