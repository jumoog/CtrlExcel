#include <ExternHdl.hxx>

#include <ExcelXlsxHelpers.hxx>

#include <DynVar.hxx>
#include <MappingVar.hxx>

#include <filesystem>
#include <fstream>
#include <string>

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

      if ( !hasNumArgs(1, param) )
        return &dynTextResult;

      TextVar filenameVar;
      filenameVar = *(param.args->getFirst()->evaluate(param.thread));

      try
      {
        XLDocument doc;
        doc.open(filenameVar.getValue());
        auto names = doc.workbook().sheetNames();
        for ( const auto &name : names )
          dynTextResult.append(TextVar(name.c_str()));
        doc.close();
      }
      catch (...) {}

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

      TextVar filenameVar, sheetnameVar;
      filenameVar  = *(param.args->getFirst()->evaluate(param.thread));
      sheetnameVar = *(param.args->getNext() ->evaluate(param.thread));

      bool skipHidden = true;
      CtrlExpr *skipArg = param.args->getNext();
      if ( skipArg )
      {
        BitVar skipHiddenVar;
        skipHiddenVar = *(skipArg->evaluate(param.thread));
        skipHidden = skipHiddenVar.isTrue();
      }

      bool useHeaders = true;
      CtrlExpr *headerArg = param.args->getNext();
      if ( headerArg )
      {
        BitVar headerVar;
        headerVar = *(headerArg->evaluate(param.thread));
        useHeaders = headerVar.isTrue();
      }

      try
      {
        XLDocument doc;
        doc.open(filenameVar.getValue());

        const char *sheetname = sheetnameVar.getValue();
        auto wks = (*sheetname)
          ? doc.workbook().worksheet(std::string(sheetname))
          : doc.workbook().worksheet(1);

        ExcelXlsxHelpers::readSheetRows(wks, doc, dynMappingResult, useHeaders, skipHidden);
        doc.close();
      }
      catch (...) {}

      return &dynMappingResult;
    }

    // -------------------------------------------------------------------------
    // excelReadFile(string filename, bool skipHiddenRows,
    //              bool firstRowIsColumnNames) -> mapping
    case F_excelReadFile:
    {
      param.thread->clearLastError();
      mappingResult = MappingVar();

      TextVar filenameVar;
      filenameVar = *(param.args->getFirst()->evaluate(param.thread));

      bool skipHidden = true;
      CtrlExpr *skipArg = param.args->getNext();
      if ( skipArg )
      {
        BitVar skipHiddenVar;
        skipHiddenVar = *(skipArg->evaluate(param.thread));
        skipHidden = skipHiddenVar.isTrue();
      }

      bool useHeaders = true;
      CtrlExpr *headerArg = param.args->getNext();
      if ( headerArg )
      {
        BitVar headerVar;
        headerVar = *(headerArg->evaluate(param.thread));
        useHeaders = headerVar.isTrue();
      }

      try
      {
        XLDocument doc;
        doc.open(filenameVar.getValue());

        auto sheetNames = doc.workbook().worksheetNames();
        for ( const auto &sn : sheetNames )
        {
          auto wks = doc.workbook().worksheet(sn);

          DynVar sheetDyn;
          sheetDyn.reset(MAPPING_VAR);
          ExcelXlsxHelpers::readSheetRows(wks, doc, sheetDyn, useHeaders, skipHidden);

          mappingResult.setAt(TextVar(sn.c_str()), sheetDyn);
        }

        doc.close();
      }
      catch (...) {}

      return &mappingResult;
    }

    // -------------------------------------------------------------------------
    // excelWriteSheet(string filename, string sheetName, dyn_anytype data)
    //   -> bool
    case F_excelWriteSheet:
    {
      param.thread->clearLastError();
      writeResult = BitVar(false);

      TextVar filenameVar, sheetnameVar;
      filenameVar  = *(param.args->getFirst()->evaluate(param.thread));
      sheetnameVar = *(param.args->getNext() ->evaluate(param.thread));

      const Variable *dataPtr = param.args->getNext()->evaluate(param.thread);
      dataPtr = ExcelXlsxHelpers::unwrapAnyOrMixed(dataPtr);
      if ( !dataPtr || !dataPtr->isDynVar() )
        return &writeResult;

      DynVar *dataVar = const_cast<DynVar *>(static_cast<const DynVar *>(dataPtr));

      try
      {
        // Check if file exists and is locked/opened by another process
        std::filesystem::path filePath(filenameVar.getValue());
        if ( std::filesystem::exists(filePath) )
        {
          // Try to open the file in write mode to check if it's accessible
          std::ofstream testWrite(filePath, std::ios::out | std::ios::app);
          if ( !testWrite.is_open() )
          {
            // File is likely opened by another process (e.g., Excel)
            writeResult = BitVar(false);
            return &writeResult;
          }
          testWrite.close();
        }

        XLDocument doc;
        doc.create(filenameVar.getValue(), XLForceOverwrite);
        doc.setProperty(XLProperty::Creator, "WinCC OA");
        doc.setProperty(XLProperty::LastModifiedBy, "WinCC OA");

        const char *sheetname = sheetnameVar.getValue();
        std::string sheetName = (*sheetname) ? sheetname : "Sheet1";

        auto wb = doc.workbook();
        wb.worksheet(1).setName(sheetName);

        auto wks = wb.worksheet(sheetName);
        bool ok = ExcelXlsxHelpers::writeSheetData(wks, *dataVar, doc);

        doc.save();
        doc.close();

        if ( ok )
          writeResult = BitVar(true);
      }
      catch (...) {}

      return &writeResult;
    }

    // -------------------------------------------------------------------------
    // excelWriteFile(string filename, mapping data) -> bool
    case F_excelWriteFile:
    {
      param.thread->clearLastError();
      writeResult = BitVar(false);

      TextVar filenameVar;
      filenameVar = *(param.args->getFirst()->evaluate(param.thread));

      MappingVar dataVar;
      dataVar = *(param.args->getNext()->evaluate(param.thread));

      unsigned int numSheets = dataVar.getNumberOfItems();
      if ( numSheets == 0 )
      {
        writeResult = BitVar(true);
        return &writeResult;
      }

      try
      {
        // Check if file exists and is locked/opened by another process
        std::filesystem::path filePath(filenameVar.getValue());
        if ( std::filesystem::exists(filePath) )
        {
          // Try to open the file in write mode to check if it's accessible
          std::ofstream testWrite(filePath, std::ios::out | std::ios::app);
          if ( !testWrite.is_open() )
          {
            // File is likely opened by another process (e.g., Excel)
            writeResult = BitVar(false);
            return &writeResult;
          }
          testWrite.close();
        }

        XLDocument doc;
        doc.create(filenameVar.getValue(), XLForceOverwrite);
        doc.setProperty(XLProperty::Creator, "WinCC OA");
        doc.setProperty(XLProperty::LastModifiedBy, "WinCC OA");

        auto wb = doc.workbook();
        bool ok = true;

        for ( unsigned int s = 0; s < numSheets; s++ )
        {
          CharString sheetName = dataVar.getKey(s)->formatValue(CharString());
          std::string sheetNameStr(sheetName.c_str());

          if ( s == 0 )
            wb.worksheet(1).setName(sheetNameStr);
          else
            wb.addWorksheet(sheetNameStr);

          auto wks = wb.worksheet(sheetNameStr);

          const Variable *sheetDataVar = dataVar.getValue(s);
          sheetDataVar = ExcelXlsxHelpers::unwrapAnyOrMixed(sheetDataVar);
          if ( !sheetDataVar || !sheetDataVar->isDynVar() )
            continue;

          DynVar *sheetData = const_cast<DynVar *>(
            static_cast<const DynVar *>(sheetDataVar));
          ok = ExcelXlsxHelpers::writeSheetData(wks, *sheetData, doc) && ok;
        }

        doc.save();
        doc.close();

        if ( ok )
          writeResult = BitVar(true);
      }
      catch (...) {}

      return &writeResult;
    }

    // -------------------------------------------------------------------------
    default:
      return &errorIntVar;
  }
}

//------------------------------------------------------------------------------
