# Builds the Rust library in rust/ (calamine + rust_xlsxwriter behind a C ABI)
# with cargo and provides it as the target ctrlexcel_xlsx. Link it with
# ctrlexcel_link_xlsx(<target>).
#
# cargo runs in rust/, so rustup picks the toolchain from
# rust/rust-toolchain.toml; it builds in <build>/cargo. Always the release
# profile (with line tables): debug builds of calamine are slow, and the
# library links the release C runtime on Windows anyway (see below).

find_program(CARGO_EXECUTABLE cargo HINTS "$ENV{CARGO_HOME}/bin" "$ENV{HOME}/.cargo/bin" "$ENV{USERPROFILE}/.cargo/bin")
if ( NOT CARGO_EXECUTABLE )
  message(FATAL_ERROR "cargo not found: install Rust (https://rustup.rs), see README.md")
endif()

# A toolchain file for cross-compiling sets the matching Rust target
# (cmake/toolchain-aarch64-linux-gnu.cmake); empty builds for the host.
# Not set(CACHE) over a toolchain file's variable: before policy CMP0126 that
# would remove it.
if ( NOT DEFINED CTRLEXCEL_RUST_TARGET )
  set(CTRLEXCEL_RUST_TARGET "" CACHE STRING "Rust target triple of the library (empty: the host)")
endif()

set(CTRLEXCEL_RUST_DIR ${PROJECT_SOURCE_DIR}/rust)
set(CTRLEXCEL_CARGO_TARGET_DIR ${PROJECT_BINARY_DIR}/cargo)
if ( CTRLEXCEL_RUST_TARGET )
  set(CTRLEXCEL_CARGO_TARGET_ARGS --target ${CTRLEXCEL_RUST_TARGET})
  set(CTRLEXCEL_CARGO_OUT_DIR ${CTRLEXCEL_CARGO_TARGET_DIR}/${CTRLEXCEL_RUST_TARGET}/release)
else()
  set(CTRLEXCEL_CARGO_TARGET_ARGS)
  set(CTRLEXCEL_CARGO_OUT_DIR ${CTRLEXCEL_CARGO_TARGET_DIR}/release)
endif()
if ( MSVC )
  set(CTRLEXCEL_RUST_LIB ${CTRLEXCEL_CARGO_OUT_DIR}/ctrlexcel_xlsx.lib)
else()
  set(CTRLEXCEL_RUST_LIB ${CTRLEXCEL_CARGO_OUT_DIR}/libctrlexcel_xlsx.a)
endif()

file(GLOB_RECURSE CTRLEXCEL_RUST_SOURCES CONFIGURE_DEPENDS ${CTRLEXCEL_RUST_DIR}/src/*.rs)
add_custom_command(
  OUTPUT ${CTRLEXCEL_RUST_LIB}
  COMMAND ${CMAKE_COMMAND} -E env CARGO_TARGET_DIR=${CTRLEXCEL_CARGO_TARGET_DIR}
          ${CARGO_EXECUTABLE} build --release --locked --lib ${CTRLEXCEL_CARGO_TARGET_ARGS}
  WORKING_DIRECTORY ${CTRLEXCEL_RUST_DIR}
  DEPENDS ${CTRLEXCEL_RUST_SOURCES}
          ${CTRLEXCEL_RUST_DIR}/Cargo.toml
          ${CTRLEXCEL_RUST_DIR}/Cargo.lock
          ${CTRLEXCEL_RUST_DIR}/rust-toolchain.toml
  COMMENT "cargo build (rust/)"
  USES_TERMINAL
  VERBATIM)
add_custom_target(ctrlexcel_xlsx_cargo DEPENDS ${CTRLEXCEL_RUST_LIB})

add_library(ctrlexcel_xlsx INTERFACE)
target_include_directories(ctrlexcel_xlsx INTERFACE ${CTRLEXCEL_RUST_DIR}/include)
target_link_libraries(ctrlexcel_xlsx INTERFACE ${CTRLEXCEL_RUST_LIB})
if ( WIN32 )
  # cargo rustc -- --print native-static-libs
  target_link_libraries(ctrlexcel_xlsx INTERFACE kernel32 ntdll userenv ws2_32 dbghelp synchronization)
  # Rust always links the release C runtime (msvcrt). WinCC OA's
  # CMakeDefines.txt uses it (/MD) in Debug builds too, and CMakeLists.txt
  # does the same without the API, so there is only one runtime.
else()
  find_package(Threads REQUIRED)
  target_link_libraries(ctrlexcel_xlsx INTERFACE Threads::Threads ${CMAKE_DL_LIBS} m)
endif()

# Links target with the Rust library (and builds that first).
function(ctrlexcel_link_xlsx target)
  add_dependencies(${target} ctrlexcel_xlsx_cargo)
  # Plain signature: add_ctrlext (CMakeDefines.txt) uses it for the extension.
  target_link_libraries(${target} ctrlexcel_xlsx)
endfunction()

# The crate versions, for the version resource.
file(READ ${CTRLEXCEL_RUST_DIR}/Cargo.lock CTRLEXCEL_CARGO_LOCK)
foreach ( crate calamine rust_xlsxwriter )
  string(REGEX MATCH "name = \"${crate}\"\r?\nversion = \"([^\"]+)\"" _match "${CTRLEXCEL_CARGO_LOCK}")
  set(CTRLEXCEL_${crate}_VERSION "${CMAKE_MATCH_1}")
endforeach()
