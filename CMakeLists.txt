cmake_minimum_required(VERSION 3.15)

# Without the (proprietary) WinCC OA API only the tests are built, e.g. in CI:
# cmake -DCTRLEXCEL_BUILD_EXTENSION=OFF
option(CTRLEXCEL_BUILD_EXTENSION "Build the CTRL extension (needs the WinCC OA API in API_ROOT)" ON)
option(CTRLEXCEL_BUILD_TESTS "Build the C++ tests (tests/)" ON)

if ( CTRLEXCEL_BUILD_EXTENSION )
  set(API_ROOT "$ENV{API_ROOT}" CACHE FILEPATH "directory of the WinCC_OA API installation")
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
  ExternHdl.cxx
  ExcelXlsxHelpers.cxx
  ExcelXlsxCore.cxx
  )

  if ( WIN32 )
    set (SOURCES ${SOURCES} VersInfo.rc)
  endif()

  add_ctrlext(${TARGET} ${SOURCES})

  target_link_libraries(${TARGET} OpenXLSX::OpenXLSX Threads::Threads)
endif()

if ( CTRLEXCEL_BUILD_TESTS )
  enable_testing()
  add_subdirectory(tests)
endif()
