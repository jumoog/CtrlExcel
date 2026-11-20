#ifndef _EXCEL_XLSX_HELPERS_HXX_
#define _EXCEL_XLSX_HELPERS_HXX_

#include <AnyTypeVar.hxx>
#include <DynVar.hxx>
#include <MixedVar.hxx>

#include <CharString.hxx>

#include <string>
#include <vector>

#include <ExcelXlsxCore.hxx>

class Variable;

namespace ExcelXlsxHelpers
{
  const Variable *unwrapAnyOrMixed(const Variable *val);

  // Conversion between the project's text encoding (what CTRL strings hold)
  // and UTF-8 (what the xlsx library and the file use). No-ops in UTF-8 projects.
  std::string toUtf8(const char *projectText);
  CharString fromUtf8(const std::string &utf8);

  // False if the UTF-8 text contains characters the project encoding lacks.
  bool isRepresentable(const std::string &utf8);

  // Appends one mapping per data row of the worksheet sheetName (UTF-8),
  // with texts and header keys in UTF-8 (safe on a worker thread; see
  // convertRowsToProjectEncoding). Rows without any value are skipped if
  // skipEmpty. warnings (UTF-8) receives non-fatal problems such as formula
  // cells without a calculated value. Throws std::runtime_error.
  void readSheetRows(ExcelXlsxCore::Workbook &book, const std::string &sheetName,
                     DynVar &result, bool useHeaders, bool skipHidden,
                     bool skipEmpty, std::vector<std::string> &warnings);

  // Converts the UTF-8 texts and text keys produced by readSheetRows to the
  // project encoding, suffixing keys that collide after conversion. Uses
  // WinCC OA's converter: CTRL thread only. No-op in UTF-8 projects.
  void convertRowsToProjectEncoding(DynVar &rows, std::vector<std::string> &warnings);

  // Writes the rows to the sheet added last to out. Returns false with a
  // reason in error if a row is not a mapping, the data exceeds Excel's
  // sheet size, or a text exceeds Excel's cell limit; out must not be saved
  // then. Throws std::runtime_error if the writer fails.
  bool writeSheetData(ExcelXlsxCore::WorkbookWriter &out, const DynVar &data, std::string &error);
}

#endif
