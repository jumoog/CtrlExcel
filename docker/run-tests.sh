#!/bin/bash
# Builds CtrlExcelReader against the WinCC OA API and runs both test suites:
# the C++ tests (ctest) and ExcelRoundTripTest.ctl in a minimal WinCC OA
# project, with a CTRL manager that does not connect to Data/Event (-n), so
# no database or PMON is needed. See README.md, "Build and test on Linux in Docker".
#
# Expects the sources at /src (read-only is fine) and writes to /build.
# Exit code 0 only if everything built and every test passed.
set -euo pipefail

SRC=${SRC:-/src}
BUILD=${BUILD:-/build}
OAINST=${OAINST:-/opt/WinCC_OA/3.21}
CTRL_TIMEOUT=${CTRL_TIMEOUT:-600}

echo "=== Configure"
cmake -S "$SRC" -B "$BUILD/cmake" -G Ninja \
  -DCMAKE_BUILD_TYPE=RelWithDebInfo \
  -DCMAKE_TOOLCHAIN_FILE="$VCPKG_ROOT/scripts/buildsystems/vcpkg.cmake" \
  -DVCPKG_TARGET_TRIPLET=x64-linux \
  -DCMAKE_COMPILE_WARNING_AS_ERROR="${WARNINGS_AS_ERRORS:-ON}"

echo "=== Build"
cmake --build "$BUILD/cmake" --parallel

echo "=== C++ tests"
ctest --test-dir "$BUILD/cmake" --output-on-failure

echo "=== CTRL tests"
PROJ="$BUILD/proj"
rm -rf "$PROJ"
mkdir -p "$PROJ/config" "$PROJ/bin" "$PROJ/scripts" "$PROJ/data" "$PROJ/log"
cat > "$PROJ/config/config" <<EOF
[general]
pvss_path = "$OAINST"
proj_path = "$PROJ"
proj_version = "3.21"
langs = "en_US.utf8"
EOF
cp "$BUILD/cmake/CtrlExcelReader.so" "$PROJ/bin/"
cp -r "$SRC/scripts/." "$PROJ/scripts/"   # libs/classes/ExcelWorkbook.ctl
cp "$SRC/ExcelRoundTripTest.ctl" "$SRC/ExcelWorkbookTest.ctl" "$PROJ/scripts/"
cp "$SRC/testdata/CtrlExcelReaderFixture.xlsx" "$PROJ/data/"

# Runs one CTRL test script; each logs one "<name> summary" line (ALL PASSED
# or FAILURES), followed by the failed and skipped test names.
run_ctrl_test()
{
  local name=$1
  local log="$BUILD/$name.log"
  local status=0
  echo "--- $name"
  # -autofreg registers the throw-away project on the fly.
  timeout "$CTRL_TIMEOUT" "$OAINST/bin/WCCOActrl" -config "$PROJ/config/config" -autofreg -n \
    -log +stderr "$name.ctl" > "$log" 2>&1 || status=$?
  cat "$log"
  if [ "$status" -eq 124 ]; then
    echo "$name timed out after ${CTRL_TIMEOUT} s" >&2
    return 1
  fi
  if ! grep -q "\"$name summary\"\]\[\"ALL PASSED\"\]" "$log"; then
    echo "$name failed (WCCOActrl exit code $status), see $log" >&2
    return 1
  fi
}

failed=0
run_ctrl_test ExcelRoundTripTest || failed=1
run_ctrl_test ExcelWorkbookTest || failed=1
[ "$failed" -eq 0 ] || exit 1
echo "=== All tests passed"
