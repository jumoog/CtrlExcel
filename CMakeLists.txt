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
endif()

project(CtrlExcelReader)

# --- OpenXLSX (Excel reader / writer) ----------------------------------------
# OpenXLSXConfig.cmake does not call find_dependency for its link deps
find_package(pugixml CONFIG REQUIRED)
find_package(miniz CONFIG REQUIRED)
if ( WIN32 )
  # The vcpkg port uses Boost.Nowide on Windows only
  find_package(Boost CONFIG REQUIRED COMPONENTS nowide)
endif()
find_package(OpenXLSX CONFIG REQUIRED)

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

  # pugixml directly too: ExcelXlsxCore walks the sheet XML (forEachRow)
  target_link_libraries(${TARGET} OpenXLSX::OpenXLSX pugixml::pugixml Threads::Threads)

  if ( NOT WIN32 )
    # Export only newExternHdl and getVersionString. The CTRL manager loads
    # extensions into one global symbol namespace, where any other exported
    # symbol (our classes, or the statically linked OpenXLSX, pugixml and
    # miniz) can bind to, or be replaced by, a same-named one of WinCC OA.
    set_target_properties(${TARGET} PROPERTIES CXX_VISIBILITY_PRESET hidden)
    target_link_options(${TARGET} PRIVATE "LINKER:--exclude-libs,ALL")
  endif()
endif()

if ( CTRLEXCEL_BUILD_TESTS )
  enable_testing()
  add_subdirectory(tests)
endif()
