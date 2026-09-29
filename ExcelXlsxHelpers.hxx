#ifndef _EXCEL_XLSX_HELPERS_HXX_
#define _EXCEL_XLSX_HELPERS_HXX_

#include <AnyTypeVar.hxx>
#include <DynVar.hxx>
#include <MixedVar.hxx>

#include <CharString.hxx>

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

  // Conversion between the project's text encoding (what CTRL strings hold)
  // and UTF-8 (what OpenXLSX and the file use). No-ops in UTF-8 projects.
  std::string toUtf8(const char *projectText);
  CharString fromUtf8(const std::string &utf8);

  // False if the UTF-8 text contains characters the project encoding lacks.
  bool isRepresentable(const std::string &utf8);

  // Appends one mapping per data row, with texts and header keys in UTF-8
  // (safe on a worker thread; see convertRowsToProjectEncoding). Rows
  // without any value are skipped if skipEmpty. warnings (UTF-8) receives
  // non-fatal problems such as formula cells without a calculated value.
  void readSheetRows(OpenXLSX::XLWorksheet &wks, OpenXLSX::XLDocument &doc,
                     DynVar &result, bool useHeaders, bool skipHidden,
                     bool skipEmpty, std::vector<std::string> &warnings);

  // Converts the UTF-8 texts and text keys produced by readSheetRows to the
  // project encoding, suffixing keys that collide after conversion. Uses
  // WinCC OA's converter: CTRL thread only. No-op in UTF-8 projects.
  void convertRowsToProjectEncoding(DynVar &rows, std::vector<std::string> &warnings);

  // Returns false with a reason in error if a row is not a mapping, the data
  // exceeds Excel's sheet size, or a text exceeds Excel's cell limit.
  bool writeSheetData(OpenXLSX::XLWorksheet &wks, const DynVar &data,
                      OpenXLSX::XLDocument &doc, std::string &error);

  // names are UTF-8 (see toUtf8); converts them to valid UTF-8 in place. Returns an empty string if
  // Excel accepts every name, otherwise the first problem found.
  std::string checkSheetNames(std::vector<std::string> &names);
}

#endif
