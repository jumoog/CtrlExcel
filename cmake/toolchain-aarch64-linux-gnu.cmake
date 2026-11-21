# Cross-compiling for Linux arm64 with Debian's cross compiler
# (g++-aarch64-linux-gnu), e.g. in the image of
# docker/winccoa-debian-arm64-cross:
#
#   cmake -DCMAKE_TOOLCHAIN_FILE=cmake/toolchain-aarch64-linux-gnu.cmake ...
#
# The Rust library is built for the same target (rustup target add
# aarch64-unknown-linux-gnu). With qemu-user installed, ctest runs the tests
# through qemu-aarch64.
set(CMAKE_SYSTEM_NAME Linux)
set(CMAKE_SYSTEM_PROCESSOR aarch64)

set(CMAKE_C_COMPILER aarch64-linux-gnu-gcc)
set(CMAKE_CXX_COMPILER aarch64-linux-gnu-g++)

# Unless given (-DCTRLEXCEL_RUST_TARGET, as docker/run-tests.sh does).
if ( NOT DEFINED CTRLEXCEL_RUST_TARGET )
  set(CTRLEXCEL_RUST_TARGET aarch64-unknown-linux-gnu)
endif()

find_program(CTRLEXCEL_QEMU_AARCH64 qemu-aarch64)
if ( CTRLEXCEL_QEMU_AARCH64 )
  set(CMAKE_CROSSCOMPILING_EMULATOR ${CTRLEXCEL_QEMU_AARCH64})
endif()
