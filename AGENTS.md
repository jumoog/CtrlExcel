# AGENTS.md

Guidance for coding agents working in this repository. User-facing API docs live in `README.md`.

## What this is

A WinCC OA CTRL extension (`CtrlExcelReader`) that reads and writes `.xlsx` files via OpenXLSX (vcpkg). It exposes five CTRL functions: `excelGetSheetNames`, `excelReadSheet`, `excelReadFile`, `excelWriteSheet` and `excelWriteFile`.

| File | Role |
|---|---|
| `ExternHdl.cxx/.hxx` | CTRL entry point: `fnList` signatures, argument evaluation, error reporting, `writeWorkbook` |
| `ExcelXlsxHelpers.cxx/.hxx` | Cell ⇄ WinCC OA `Variable` conversion, header keys, date handling, `readSheetRows` / `writeSheetData` |
| `ExcelRoundTripTest.ctl` | CTRL round-trip tests (need a WinCC OA project to run) |
| `CMakeLists.txt`, `CMakePresets.json`, `vcpkg.json` | Build |

## Build

Requires `API_ROOT` (WinCC OA `api/` dir) and `VCPKG_ROOT` in the environment.

```powershell
cmake --preset vs2022-vcpkg
cmake --build --preset relwithdebinfo
```

Output: `build/RelWithDebInfo/CtrlExcelReader.dll`. Linux presets: `linux-vcpkg-relwithdebinfo` / `linux-relwithdebinfo`. C++17 is set by the WinCC OA `CMakeDefines.txt`.

Always build after changing C++ code and keep it warning-free.

## Testing

There is no C++ test harness. `ExcelRoundTripTest.ctl` runs inside a WinCC OA project with the built DLL loaded (`LoadCtrlLibs = "CtrlExcelReader"`); each test logs `pass` via `DebugTN`. An agent cannot run it: say so and ask the user to run it after behaviour changes. After every rebuild, the new DLL must be copied into the project's `bin/` and the CTRL manager restarted (a running manager keeps the old DLL loaded). A test failure that contradicts the current code is most likely a stale DLL: ask for the DLL's timestamp before debugging. Add a test function there (and call it from `main()`) for new behaviour that the write API can produce.

## Known pitfalls

- **`ssize_t` clash**: WinCC OA's `winnt/win32.h` typedefs `ssize_t` as `int`; OpenXLSX redefines it as 64-bit. Include OpenXLSX only through `ExcelXlsxHelpers.hxx`, which renames it via `#define ssize_t OpenXLSX_ssize_t`. Never `#include <OpenXLSX.hpp>` directly.
- **OpenXLSX CMake config** does not `find_dependency` its link deps; `CMakeLists.txt` must `find_package` pugixml, miniz and Boost (nowide) before OpenXLSX.
- **`XLDocument::create()`** builds in a temp archive; the target file is only written by `save()`. `writeWorkbook` relies on this: validate everything first, and skip `save()` on failure so an existing file is never clobbered.
- **Dates**: `XLDateTime::tm()` truncates to whole seconds, and `XLDateTime(serial)` throws for serial < 1. Use `serialToLocalTime()` (rounds the whole serial to ms once) and `isDateValue()` (excludes serial < 1). Excel serials are local wall-clock time without a timezone; the DST fall-back hour is inherently ambiguous (documented in README).
- **Header keys** (`readSheetRows`): text for non-empty headers, repeated names suffixed `_2`, `_3`…, empty headers fall back to the 1-based column number as `int`. Numeric and date headers go through `headerText()`; never `get<std::string>()` on a non-string cell (it throws).
- **Writing**: look up row values with the first row's original key `Variable`s, not stringified `TextVar`s, or non-text keys (e.g. int column keys) miss.
- **Integers**: values outside the 32-bit range must become `LongVar`, not a truncated `IntegerVar`.
- **Filenames** from CTRL are UTF-8; build `std::filesystem` paths with `std::filesystem::u8path`.

## CTRL extension conventions

- `ExecuteParamRec` is nested in `BaseExternHdl`; helpers outside the class take `CtrlThread *`.
- Every function starts with `param.thread->clearLastError()`, then `hasNumArgs(min, max, param)`.
- Evaluate arguments with `evalArg` / `evalOptionalArg`: `evaluate()` may return `nullptr`.
- Never swallow exceptions silently: use `reportError` / `reportCurrentException` so scripts see them via `getLastError()`. Include the file (and sheet) name in the context.
- Results are returned via function-local `static` variables, reset at the start of each call.
- Keep `fnList` signatures in sync with the README.

## Style

- Match the surrounding file: `ExternHdl.cxx` uses `if ( cond )` with inner spaces; `ExcelXlsxHelpers.cxx` uses `if (cond)`. Two-space indent, braces on their own line.
- Comments explain *why* (library quirks, invariants), not what.
- Update `README.md` when user-visible behaviour changes.

## Git

- Conventional commits: `type(scope): summary`, e.g. `fix(excel): ...`, `refactor(excel): ...`, `test(excel): ...`, `docs: ...`, `fix(build): ...`.
- `build/` and `vcpkg_installed/` are ignored; never commit build output.
