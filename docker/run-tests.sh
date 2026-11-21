#!/bin/bash
# Builds CtrlExcelReader against the WinCC OA API and runs all test suites:
# the Rust tests (cargo test), the C++ tests (ctest) and the CTRL tests in a
# minimal WinCC OA project, with a CTRL manager that does not connect to
# Data/Event (-n), so no database or PMON is needed. See README.md, "Build
# and test on Linux in Docker".
#
# Expects the sources at /src (read-only is fine) and writes to /build.
# Exit code 0 only if everything built and every test passed.
set -euo pipefail

SRC=${SRC:-/src}
BUILD=${BUILD:-/build}
OAINST=${OAINST:-/opt/WinCC_OA/3.21}
CTRL_TIMEOUT=${CTRL_TIMEOUT:-600}
# Cross-compiling (set by docker/winccoa-debian-arm64-cross): the Rust target,
# the toolchain file in $SRC/cmake and the emulator for WCCOActrl; cargo test
# and ctest run their tests through the emulator, too
# (CARGO_TARGET_<triple>_RUNNER, CMAKE_CROSSCOMPILING_EMULATOR).
RUST_TARGET=${CTRLEXCEL_RUST_TARGET:-}
TOOLCHAIN=${CTRLEXCEL_TOOLCHAIN:+$SRC/cmake/$CTRLEXCEL_TOOLCHAIN}
CTRL_RUNNER=${CTRLEXCEL_CTRL_RUNNER:-}

if [ "${WARNINGS_AS_ERRORS:-ON}" = ON ]; then
  export RUSTFLAGS="${RUSTFLAGS:-} -D warnings"
fi

echo "=== Rust toolchain"
# The one pinned in rust/rust-toolchain.toml; kept in ~/.rustup (a volume).
(cd "$SRC/rust" && rustup toolchain install && cargo --version)
if [ -n "$RUST_TARGET" ]; then
  (cd "$SRC/rust" && rustup target add "$RUST_TARGET")
fi

echo "=== Rust tests"
# /src is read-only: cargo writes its target directory to /build.
(cd "$SRC/rust" && CARGO_TARGET_DIR="$BUILD/cargo-test" cargo test --locked ${RUST_TARGET:+--target "$RUST_TARGET"})

echo "=== Configure"
# --fresh: a cache from an older build setup must not carry over.
cmake --fresh -S "$SRC" -B "$BUILD/cmake" -G Ninja \
  -DCMAKE_BUILD_TYPE=RelWithDebInfo \
  ${TOOLCHAIN:+-DCMAKE_TOOLCHAIN_FILE="$TOOLCHAIN" -DCTRLEXCEL_RUST_TARGET="$RUST_TARGET"} \
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
  timeout "$CTRL_TIMEOUT" $CTRL_RUNNER "$OAINST/bin/WCCOActrl" -config "$PROJ/config/config" -autofreg -n \
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
