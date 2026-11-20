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

# The version is kept in vcpkg.json only: its top-level "version". CMake
# before 3.19 has no string(JSON) and falls back to the first "version" key.
file(READ "${CMAKE_CURRENT_SOURCE_DIR}/vcpkg.json" CTRLEXCEL_MANIFEST)
if ( CMAKE_VERSION VERSION_GREATER_EQUAL 3.19 )
  string(JSON CTRLEXCEL_VERSION ERROR_VARIABLE CTRLEXCEL_VERSION_ERROR GET "${CTRLEXCEL_MANIFEST}" version)
else()
  string(REGEX MATCH "\"version\"[ \t\r\n]*:[ \t\r\n]*\"([^\"]*)\"" CTRLEXCEL_VERSION_MATCH "${CTRLEXCEL_MANIFEST}")
  set(CTRLEXCEL_VERSION "${CMAKE_MATCH_1}")
endif()
if ( NOT CTRLEXCEL_VERSION MATCHES "^[0-9]+\\.[0-9]+\\.[0-9]+$" )
  message(FATAL_ERROR "vcpkg.json: \"version\" must be \"major.minor.patch\", got \"${CTRLEXCEL_VERSION}\"")
endif()

project(CtrlExcelReader VERSION ${CTRLEXCEL_VERSION})

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
              -DOPENXLSX_VERSION=${OpenXLSX_VERSION}
              -P ${PROJECT_SOURCE_DIR}/cmake/VersionInfo.cmake
      BYPRODUCTS ${CTRLEXCEL_VERSION_HEADER}
      COMMENT "Version information for VersInfo.rc"
      VERBATIM)
    add_dependencies(${TARGET} CtrlExcelVersion)
    target_include_directories(${TARGET} PRIVATE ${PROJECT_BINARY_DIR})
  endif()

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
