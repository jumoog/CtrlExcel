cmake_minimum_required(VERSION 3.15)

set(API_ROOT "$ENV{API_ROOT}" CACHE FILEPATH "directory of the WinCC_OA API installation")
include(${API_ROOT}/CMakeDefines.txt)

project(CtrlExcelReader)

set(TARGET CtrlExcelReader)

set(SOURCES
ExternHdl.cxx
ExcelXlsxHelpers.cxx
)

if ( WIN32 )
  set (SOURCES ${SOURCES} VersInfo.rc)
endif()

add_ctrlext(${TARGET} ${SOURCES})

# --- OpenXLSX (Excel reader / writer) ----------------------------------------
# OpenXLSXConfig.cmake does not call find_dependency for its link deps
find_package(pugixml CONFIG REQUIRED)
find_package(miniz CONFIG REQUIRED)
find_package(Boost CONFIG REQUIRED COMPONENTS nowide)
find_package(OpenXLSX CONFIG REQUIRED)
target_link_libraries(${TARGET} OpenXLSX::OpenXLSX)
