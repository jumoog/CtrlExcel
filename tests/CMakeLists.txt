# Tests of ExcelXlsxCore: everything that needs the xlsx library (rust/) but
# not the WinCC OA API. The CTRL-level behaviour is tested by
# ExcelRoundTripTest.ctl; the Rust library has its own tests (cargo test).

add_executable(ExcelXlsxCoreTest
  ExcelXlsxCoreTest.cxx
  ${PROJECT_SOURCE_DIR}/ExcelXlsxCore.cxx
)

set(CTRLEXCEL_FIXTURE ${PROJECT_SOURCE_DIR}/testdata/CtrlExcelReaderFixture.xlsx)
configure_file(FixturePath.hxx.in FixturePath.hxx @ONLY)

target_include_directories(ExcelXlsxCoreTest PRIVATE ${PROJECT_SOURCE_DIR} ${CMAKE_CURRENT_BINARY_DIR})
ctrlexcel_link_xlsx(ExcelXlsxCoreTest)

if ( MSVC )
  target_compile_options(ExcelXlsxCoreTest PRIVATE /W4 /utf-8)
else()
  target_compile_options(ExcelXlsxCoreTest PRIVATE -Wall -Wextra)
endif()

add_test(NAME ExcelXlsxCoreTest COMMAND ExcelXlsxCoreTest)
