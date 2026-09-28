#include <ExternHdl.hxx>

#include <ExcelXlsxHelpers.hxx>

#include <DynVar.hxx>
#include <ErrClass.hxx>
#include <MappingVar.hxx>

#include <exception>
#include <filesystem>
#include <fstream>
#include <stdexcept>
#include <string>
#include <utility>
#include <vector>

using namespace OpenXLSX;

static FunctionListRec fnList[] =
{
  { DYNTEXT_VAR,       "excelGetSheetNames", "(string filename)",                                                                          false },
  { DYNMAPPING_VAR,    "excelReadSheet",     "(string filename, string sheetName, bool skipHiddenRows = true, bool firstRowIsColumnNames = true)", false },
  { MAPPING_VAR,       "excelReadFile",      "(string filename, bool skipHiddenRows = true, bool firstRowIsColumnNames = true)",                   false },
  { BIT_VAR,           "excelWriteSheet",    "(string filename, string sheetName, dyn_anytype data)",                                              false },
  { BIT_VAR,           "excelWriteFile",     "(string filename, mapping data)",                                                                    false },
};

CTRL_EXTENSION(ExternHdl, fnList)

//------------------------------------------------------------------------------

namespace
{
  // Append an error the CTRL script can retrieve with getLastError().
  void reportError(CtrlThread *thread, const char *funcName,
                   const std::string &msg)
  {
    ErrClass err(ErrClass::PRIO_WARNING, ErrClass::ERR_CONTROL,
                 ErrClass::UNEXPECTEDSTATE, "CtrlExcelReader", funcName,
                 msg.c_str());
    thread->appendLastError(&err);
  }

  // Report the exception currently being handled; call only from a catch block.
  void reportCurrentException(CtrlThread *thread, const char *funcName,
                              const std::string &context = std::string())
  {
    std::string prefix = context.empty() ? std::string() : context + ": ";
    try
    {
      throw;
    }
    catch (const std::exception &e)
    {
      reportError(thread, funcName, prefix + e.what());
    }
    catch (...)
    {
      reportError(thread, funcName, prefix + "unknown error");
    }
  }

  // Evaluate expr into out. False if the argument is missing or its
  // evaluation failed (evaluate() then returns nullptr).
  template <typename VarT>
  bool evalArg(CtrlExpr *expr, CtrlThread *thread, VarT &out)
  {
    const Variable *value = expr ? expr->evaluate(thread) : nullptr;
    if ( !value )
      return false;

    out = *value;
    return true;
  }

  // Like evalArg, but an omitted argument keeps out's default value.
  template <typename VarT>
  bool evalOptionalArg(CtrlExpr *expr, CtrlThread *thread, VarT &out)
  {
    return !expr || evalArg(expr, thread, out);
  }

  // False if the file exists but cannot be opened for writing, e.g. because
  // Excel holds it open. filename is UTF-8, hence u8path.
  bool isWritableOrMissing(const std::string &filename)
  {
    std::filesystem::path filePath = std::filesystem::u8path(filename);
    if ( !std::filesystem::exists(filePath) )
      return true;

    std::ofstream testWrite(filePath, std::ios::out | std::ios::app);
    return testWrite.is_open();
  }

  // Sheet name (UTF-8) and its rows (a dyn of mappings, possibly
  // anytype-wrapped).
  using SheetList = std::vector<std::pair<std::string, const Variable *>>;

  // Filename or sheet name from CTRL, converted for OpenXLSX (UTF-8).
  std::string utf8Arg(const TextVar &var)
  {
    return ExcelXlsxHelpers::toUtf8(var.getValue());
  }

  // Name from the file, converted for CTRL (project encoding).
  TextVar projectText(const std::string &utf8)
  {
    return TextVar(ExcelXlsxHelpers::fromUtf8(utf8).c_str());
  }

  // Write sheets to a new workbook at filename (UTF-8). The file is only
  // saved when every sheet is valid, so a failed write leaves an existing
  // file untouched.
  bool writeWorkbook(CtrlThread *thread, const char *funcName,
                     const std::string &filename, const SheetList &sheets)
  {
    // OpenXLSX only rejects exact duplicates; names Excel refuses would give
    // a file Excel reports as corrupt.
    // checkSheetNames also converts the names to valid UTF-8; use its result.
    std::vector<std::string> names;
    names.reserve(sheets.size());
    for ( const auto &sheet : sheets )
      names.push_back(sheet.first);

    std::string nameProblem = ExcelXlsxHelpers::checkSheetNames(names);
    if ( !nameProblem.empty() )
    {
      reportError(thread, funcName, nameProblem);
      return false;
    }

    std::vector<const DynVar *> sheetRows;
    sheetRows.reserve(sheets.size());
    for ( const auto &sheet : sheets )
    {
      const Variable *rows = ExcelXlsxHelpers::unwrapAnyOrMixed(sheet.second);
      if ( !rows || !rows->isDynVar() )
      {
        reportError(thread, funcName, sheet.first + ": data must be a dyn_mapping");
        return false;
      }
      sheetRows.push_back(static_cast<const DynVar *>(rows));
    }

    try
    {
      if ( !isWritableOrMissing(filename) )
      {
        reportError(thread, funcName, filename + ": file is locked or not writable");
        return false;
      }

      // create() works on a temporary archive; filename is only written by save().
      XLDocument doc;
      doc.create(filename, XLForceOverwrite);
      doc.setProperty(XLProperty::Creator, "WinCC OA");
      doc.setProperty(XLProperty::LastModifiedBy, "WinCC OA");

      auto wb = doc.workbook();
      for ( size_t s = 0; s < sheets.size(); s++ )
      {
        const std::string &sheetName = names[s];
        if ( s == 0 )
          wb.worksheet(1).setName(sheetName);
        else
          wb.addWorksheet(sheetName);

        auto wks = wb.worksheet(sheetName);
        std::string error;
        if ( !ExcelXlsxHelpers::writeSheetData(wks, *sheetRows[s], doc, error) )
        {
          reportError(thread, funcName, sheetName + ": " + error);
          return false;
        }
      }

      doc.save();
      doc.close();
      return true;
    }
    catch (...)
    {
      reportCurrentException(thread, funcName, filename);
      return false;
    }
  }
}

//------------------------------------------------------------------------------

const Variable *ExternHdl::execute(ExecuteParamRec &param)
{
  enum
  {
    F_excelGetSheetNames = 0,
    F_excelReadSheet     = 1,
    F_excelReadFile      = 2,
    F_excelWriteSheet    = 3,
    F_excelWriteFile     = 4,
  };

  static DynVar dynTextResult;
  static DynVar dynMappingResult;
  static MappingVar mappingResult;
  static BitVar writeResult;

  switch ( param.funcNum )
  {
    // -------------------------------------------------------------------------
    // excelGetSheetNames(string filename) -> dyn_string
    case F_excelGetSheetNames:
    {
      param.thread->clearLastError();
      dynTextResult.reset(TEXT_VAR);

      if ( !hasNumArgs(1, 1, param) )
        return &dynTextResult;

      TextVar filenameVar;
      if ( !evalArg(param.args->getFirst(), param.thread, filenameVar) )
      {
        reportError(param.thread, "excelGetSheetNames", "argument could not be evaluated");
        return &dynTextResult;
      }

      try
      {
        XLDocument doc;
        doc.open(utf8Arg(filenameVar));
        // Worksheets only: chartsheets cannot be read by excelReadSheet.
        auto names = doc.workbook().worksheetNames();
        for ( const auto &name : names )
          dynTextResult.append(projectText(name));
        doc.close();
      }
      catch (...)
      {
        reportCurrentException(param.thread, "excelGetSheetNames", filenameVar.getValue());
      }

      return &dynTextResult;
    }

    // -------------------------------------------------------------------------
    // excelReadSheet(string filename, string sheetName,
    //                bool skipHiddenRows, bool firstRowIsColumnNames)
    //   -> dyn_mapping
    case F_excelReadSheet:
    {
      param.thread->clearLastError();
      dynMappingResult.reset(MAPPING_VAR);

      if ( !hasNumArgs(2, 4, param) )
        return &dynMappingResult;

      TextVar filenameVar, sheetnameVar;
      BitVar skipHiddenVar(true), headerVar(true);
      if ( !evalArg(param.args->getFirst(), param.thread, filenameVar)
        || !evalArg(param.args->getNext(), param.thread, sheetnameVar)
        || !evalOptionalArg(param.args->getNext(), param.thread, skipHiddenVar)
        || !evalOptionalArg(param.args->getNext(), param.thread, headerVar) )
      {
        reportError(param.thread, "excelReadSheet", "argument could not be evaluated");
        return &dynMappingResult;
      }

      try
      {
        XLDocument doc;
        doc.open(utf8Arg(filenameVar));

        // An empty name means the first worksheet; worksheet(1) would be the
        // first sheet of any kind and fail for a chart sheet.
        std::string sheetname = utf8Arg(sheetnameVar);
        if ( sheetname.empty() )
        {
          auto names = doc.workbook().worksheetNames();
          if ( names.empty() )
            throw std::runtime_error("workbook has no worksheets");
          sheetname = names.front();
        }
        auto wks = doc.workbook().worksheet(sheetname);

        ExcelXlsxHelpers::readSheetRows(wks, doc, dynMappingResult,
                                        headerVar.isTrue(), skipHiddenVar.isTrue());
        doc.close();
      }
      catch (...)
      {
        reportCurrentException(param.thread, "excelReadSheet",
          std::string(filenameVar.getValue()) + " [" + sheetnameVar.getValue() + "]");
      }

      return &dynMappingResult;
    }

    // -------------------------------------------------------------------------
    // excelReadFile(string filename, bool skipHiddenRows,
    //              bool firstRowIsColumnNames) -> mapping
    case F_excelReadFile:
    {
      param.thread->clearLastError();
      mappingResult = MappingVar();

      if ( !hasNumArgs(1, 3, param) )
        return &mappingResult;

      TextVar filenameVar;
      BitVar skipHiddenVar(true), headerVar(true);
      if ( !evalArg(param.args->getFirst(), param.thread, filenameVar)
        || !evalOptionalArg(param.args->getNext(), param.thread, skipHiddenVar)
        || !evalOptionalArg(param.args->getNext(), param.thread, headerVar) )
      {
        reportError(param.thread, "excelReadFile", "argument could not be evaluated");
        return &mappingResult;
      }

      try
      {
        XLDocument doc;
        doc.open(utf8Arg(filenameVar));

        auto sheetNames = doc.workbook().worksheetNames();
        for ( const auto &sn : sheetNames )
        {
          // A broken sheet must not discard the sheets read so far or after it.
          try
          {
            auto wks = doc.workbook().worksheet(sn);

            DynVar sheetDyn;
            sheetDyn.reset(MAPPING_VAR);
            ExcelXlsxHelpers::readSheetRows(wks, doc, sheetDyn,
                                            headerVar.isTrue(), skipHiddenVar.isTrue());

            mappingResult.setAt(projectText(sn), sheetDyn);
          }
          catch (...)
          {
            reportCurrentException(param.thread, "excelReadFile",
              std::string(filenameVar.getValue()) + " [" + sn + "]");
          }
        }

        doc.close();
      }
      catch (...)
      {
        reportCurrentException(param.thread, "excelReadFile", filenameVar.getValue());
      }

      return &mappingResult;
    }

    // -------------------------------------------------------------------------
    // excelWriteSheet(string filename, string sheetName, dyn_anytype data)
    //   -> bool
    case F_excelWriteSheet:
    {
      param.thread->clearLastError();
      writeResult = BitVar(false);

      if ( !hasNumArgs(3, 3, param) )
        return &writeResult;

      TextVar filenameVar, sheetnameVar;
      if ( !evalArg(param.args->getFirst(), param.thread, filenameVar)
        || !evalArg(param.args->getNext(), param.thread, sheetnameVar) )
      {
        reportError(param.thread, "excelWriteSheet", "argument could not be evaluated");
        return &writeResult;
      }

      // Evaluated last and not copied; writeWorkbook reports nullptr or a
      // non-dyn value as invalid data.
      CtrlExpr *dataArg = param.args->getNext();
      const Variable *dataPtr = dataArg ? dataArg->evaluate(param.thread) : nullptr;

      std::string sheetname = utf8Arg(sheetnameVar);
      SheetList sheets{ { sheetname.empty() ? "Sheet1" : sheetname, dataPtr } };

      writeResult = BitVar(writeWorkbook(param.thread, "excelWriteSheet",
                                         utf8Arg(filenameVar), sheets));
      return &writeResult;
    }

    // -------------------------------------------------------------------------
    // excelWriteFile(string filename, mapping data) -> bool
    case F_excelWriteFile:
    {
      param.thread->clearLastError();
      writeResult = BitVar(false);

      if ( !hasNumArgs(2, 2, param) )
        return &writeResult;

      TextVar filenameVar;
      if ( !evalArg(param.args->getFirst(), param.thread, filenameVar) )
      {
        reportError(param.thread, "excelWriteFile", "argument could not be evaluated");
        return &writeResult;
      }

      // Checked instead of copied into a MappingVar: that conversion turns
      // any other type (e.g. a dyn_mapping meant for excelWriteSheet) into an
      // empty mapping, which would "succeed" without writing anything.
      CtrlExpr *dataArg = param.args->getNext();
      const Variable *dataPtr = ExcelXlsxHelpers::unwrapAnyOrMixed(
        dataArg ? dataArg->evaluate(param.thread) : nullptr);
      if ( !dataPtr || dataPtr->isA() != MAPPING_VAR )
      {
        reportError(param.thread, "excelWriteFile",
          "data must be a mapping of sheet name to dyn_mapping");
        return &writeResult;
      }

      const MappingVar &dataVar = *static_cast<const MappingVar *>(dataPtr);
      unsigned int numSheets = dataVar.getNumberOfItems();
      if ( numSheets == 0 )
      {
        writeResult = BitVar(true);
        return &writeResult;
      }

      SheetList sheets;
      sheets.reserve(numSheets);
      for ( unsigned int s = 0; s < numSheets; s++ )
      {
        CharString sheetName = dataVar.getKey(s)->formatValue(CharString());
        sheets.emplace_back(ExcelXlsxHelpers::toUtf8(sheetName.c_str()), dataVar.getValue(s));
      }

      writeResult = BitVar(writeWorkbook(param.thread, "excelWriteFile",
                                         utf8Arg(filenameVar), sheets));
      return &writeResult;
    }

    // -------------------------------------------------------------------------
    default:
      return &errorIntVar;
  }
}

//------------------------------------------------------------------------------
