#ifndef _EXCEL_XLSX_HELPERS_HXX_
#define _EXCEL_XLSX_HELPERS_HXX_

#include <AnyTypeVar.hxx>
#include <DynVar.hxx>
#include <MixedVar.hxx>

// WinCC OA's win32.h typedefs ssize_t as int; OpenXLSX redefines it as a
// 64-bit alias. Rename OpenXLSX's alias (no effect on mangled names).
#define ssize_t OpenXLSX_ssize_t
#include <OpenXLSX.hpp>
#undef ssize_t

class Variable;

namespace ExcelXlsxHelpers
{
  const Variable *unwrapAnyOrMixed(const Variable *val);

  void readSheetRows(OpenXLSX::XLWorksheet &wks, OpenXLSX::XLDocument &doc,
                     DynVar &result, bool useHeaders, bool skipHidden);

  // Returns false, leaving the sheet untouched, if any row is not a mapping.
  bool writeSheetData(OpenXLSX::XLWorksheet &wks, const DynVar &data,
                      OpenXLSX::XLDocument &doc);
}

#endif
