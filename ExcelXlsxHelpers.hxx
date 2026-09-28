#ifndef _EXCEL_XLSX_HELPERS_HXX_
#define _EXCEL_XLSX_HELPERS_HXX_

#include <AnyTypeVar.hxx>
#include <DynVar.hxx>
#include <MixedVar.hxx>

#include <string>
#include <vector>

// WinCC OA's win32.h typedefs ssize_t as int; OpenXLSX redefines it as a
// 64-bit alias. Rename OpenXLSX's alias (no effect on mangled names). Windows
// only: elsewhere there is no clash, and the macro could rename a system
// header's ssize_t first included through OpenXLSX.
#ifdef _WIN32
#define ssize_t OpenXLSX_ssize_t
#endif
#include <OpenXLSX.hpp>
#ifdef _WIN32
#undef ssize_t
#endif

class Variable;

namespace ExcelXlsxHelpers
{
  const Variable *unwrapAnyOrMixed(const Variable *val);

  void readSheetRows(OpenXLSX::XLWorksheet &wks, OpenXLSX::XLDocument &doc,
                     DynVar &result, bool useHeaders, bool skipHidden);

  // Returns false with a reason in error if a row is not a mapping or a text
  // exceeds Excel's cell limit. Rows are validated before anything is written.
  bool writeSheetData(OpenXLSX::XLWorksheet &wks, const DynVar &data,
                      OpenXLSX::XLDocument &doc, std::string &error);

  // Empty if Excel accepts every name, otherwise the first problem found.
  std::string checkSheetNames(const std::vector<std::string> &names);
}

#endif
