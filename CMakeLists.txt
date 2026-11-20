cmake_minimum_required(VERSION 3.15)

# Without the (proprietary) WinCC OA API only the tests are built, e.g. in CI:
# cmake -DCTRLEXCEL_BUILD_EXTENSION=OFF
option(CTRLEXCEL_BUILD_EXTENSION "Build the CTRL extension (needs the WinCC OA API in API_ROOT)" ON)
option(CTRLEXCEL_BUILD_TESTS "Build the C++ tests (tests/)" ON)

if ( CTRLEXCEL_BUILD_EXTENSION )
  set(API_ROOT "$ENV{API_ROOT}" CACHE FILEPATH "directory of the WinCC_OA API installation")
  # Linux: with VISIBILITY_HIDDEN, CMakeDefines marks the WinCC OA API
  # (DLLEXP_*, incl. CTRLEXT_EXPORT of the entry points) visible, so the
  # extension itself can be built with hidden visibility (see below).
  set(VISIBILITY_HIDDEN ON CACHE BOOL "compile with hidden symbol visibility" FORCE)
  include(${API_ROOT}/CMakeDefines.txt)
else()
  set(CMAKE_CXX_STANDARD 17)
  set(CMAKE_CXX_STANDARD_REQUIRED ON)
  # The release C runtime in every configuration, as WinCC OA's
  # CMakeDefines.txt sets it: the Rust library links it, too.
  set(CMAKE_MSVC_RUNTIME_LIBRARY MultiThreadedDLL)
endif()

# The version of the extension (the version resource takes it from here).
project(CtrlExcelReader VERSION 2.0.0)

# --- xlsx library: calamine (read) and rust_xlsxwriter (write), in rust/ ------
include(${PROJECT_SOURCE_DIR}/cmake/RustLibrary.cmake)

if ( CTRLEXCEL_BUILD_EXTENSION )
  # Worker threads of the *Async functions (std::thread needs -pthread on Linux)
  find_package(Threads REQUIRED)

  set(TARGET CtrlExcelReader)

  set(SOURCES
  ExcelExternHdl.cxx
  ExcelXlsxHelpers.cxx
  ExcelXlsxCore.cxx
  )

  if ( WIN32 )
    set (SOURCES ${SOURCES} VersInfo.rc)
  endif()

  add_ctrlext(${TARGET} ${SOURCES})

  if ( WIN32 )
    # CtrlExcelVersion.h for VersInfo.rc, regenerated on every build so the
    # DLL's ProductVersion names the commit it was built from.
    find_package(Git QUIET)
    set(CTRLEXCEL_VERSION_HEADER ${PROJECT_BINARY_DIR}/CtrlExcelVersion.h)
    add_custom_target(CtrlExcelVersion
      COMMAND ${CMAKE_COMMAND}
              -DOUTPUT=${CTRLEXCEL_VERSION_HEADER}
              -DSOURCE_DIR=${PROJECT_SOURCE_DIR}
              -DGIT_EXECUTABLE=${GIT_EXECUTABLE}
              -DVERSION=${PROJECT_VERSION}
              -DWCCOA_VERSION=${WCCOA_VERSION_DISP}
              -DCALAMINE_VERSION=${CTRLEXCEL_calamine_VERSION}
              -DXLSXWRITER_VERSION=${CTRLEXCEL_rust_xlsxwriter_VERSION}
              -P ${PROJECT_SOURCE_DIR}/cmake/VersionInfo.cmake
      BYPRODUCTS ${CTRLEXCEL_VERSION_HEADER}
      COMMENT "Version information for VersInfo.rc"
      VERBATIM)
    add_dependencies(${TARGET} CtrlExcelVersion)
    target_include_directories(${TARGET} PRIVATE ${PROJECT_BINARY_DIR})
  endif()

  target_link_libraries(${TARGET} Threads::Threads)
  ctrlexcel_link_xlsx(${TARGET})

  if ( NOT WIN32 )
    # Export only newExternHdl and getVersionString. The CTRL manager loads
    # extensions into one global symbol namespace, where any other exported
    # symbol (our classes, or the statically linked Rust library) can bind
    # to, or be replaced by, a same-named one of WinCC OA.
    set_target_properties(${TARGET} PROPERTIES CXX_VISIBILITY_PRESET hidden)
    target_link_options(${TARGET} PRIVATE "LINKER:--exclude-libs,ALL")
  endif()
endif()

if ( CTRLEXCEL_BUILD_TESTS )
  enable_testing()
  add_subdirectory(tests)
endif()
