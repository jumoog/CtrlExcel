# Tests of ExcelXlsxCore: everything that needs OpenXLSX but not the WinCC OA
# API. The CTRL-level behaviour is tested by ExcelRoundTripTest.ctl.

add_executable(ExcelXlsxCoreTest
  ExcelXlsxCoreTest.cxx
  ${PROJECT_SOURCE_DIR}/ExcelXlsxCore.cxx
)

set(CTRLEXCEL_FIXTURE ${PROJECT_SOURCE_DIR}/testdata/CtrlExcelReaderFixture.xlsx)
configure_file(FixturePath.hxx.in FixturePath.hxx @ONLY)

target_include_directories(ExcelXlsxCoreTest PRIVATE ${PROJECT_SOURCE_DIR} ${CMAKE_CURRENT_BINARY_DIR})
target_link_libraries(ExcelXlsxCoreTest PRIVATE OpenXLSX::OpenXLSX)

if ( MSVC )
  target_compile_options(ExcelXlsxCoreTest PRIVATE /W4 /utf-8)
else()
  target_compile_options(ExcelXlsxCoreTest PRIVATE -Wall -Wextra)
endif()

add_test(NAME ExcelXlsxCoreTest COMMAND ExcelXlsxCoreTest)
