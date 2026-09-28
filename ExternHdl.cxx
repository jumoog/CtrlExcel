#include <ExternHdl.hxx>

#include <ExcelXlsxHelpers.hxx>

#include <BitVar.hxx>
#include <DynVar.hxx>
#include <ErrClass.hxx>
#include <IntegerVar.hxx>
#include <MappingVar.hxx>
#include <TimeVar.hxx>
#include <WaitCond.hxx>

#include <atomic>
#include <exception>
#include <filesystem>
#include <fstream>
#include <functional>
#include <memory>
#include <stdexcept>
#include <string>
#include <thread>
#include <unordered_set>
#include <utility>
#include <vector>

using namespace OpenXLSX;

static FunctionListRec fnList[] =
{
  { DYNTEXT_VAR,       "excelGetSheetNames",      "(string filename)",                                                                          false },
  { DYNMAPPING_VAR,    "excelReadSheet",          "(string filename, string sheetName, bool skipHiddenRows = true, bool firstRowIsColumnNames = true, bool skipEmptyRows = true)", false },
  { MAPPING_VAR,       "excelReadFile",           "(string filename, bool skipHiddenRows = true, bool firstRowIsColumnNames = true, bool skipEmptyRows = true)",                   false },
  { BIT_VAR,           "excelWriteSheet",         "(string filename, string sheetName, dyn_anytype data)",                                              false },
  { BIT_VAR,           "excelWriteFile",          "(string filename, mapping data)",                                                                    false },
  { INTEGER_VAR,       "excelGetSheetNamesAsync", "(string filename, dyn_string &names)",                                                               false },
  { INTEGER_VAR,       "excelReadSheetAsync",     "(string filename, string sheetName, dyn_mapping &rows, bool skipHiddenRows = true, bool firstRowIsColumnNames = true, bool skipEmptyRows = true)", false },
  { INTEGER_VAR,       "excelReadFileAsync",      "(string filename, mapping &sheets, bool skipHiddenRows = true, bool firstRowIsColumnNames = true, bool skipEmptyRows = true)",                   false },
  { INTEGER_VAR,       "excelWriteSheetAsync",    "(string filename, string sheetName, dyn_anytype data, bool &ok)",                                    false },
  { INTEGER_VAR,       "excelWriteFileAsync",     "(string filename, mapping data, bool &ok)",                                                          false },
};

CTRL_EXTENSION(ExternHdl, fnList)

//------------------------------------------------------------------------------

namespace
{
  // Append an error the CTRL script can retrieve with getLastError(). msg is
  // UTF-8 (like everything from OpenXLSX) and converted to the project
  // encoding here, so callers must not mix in project-encoded text.
  // CTRL thread only.
  void reportError(CtrlThread *thread, const char *funcName,
                   const std::string &msg)
  {
    ErrClass err(ErrClass::PRIO_WARNING, ErrClass::ERR_CONTROL,
                 ErrClass::UNEXPECTEDSTATE, "CtrlExcelReader", funcName,
                 ExcelXlsxHelpers::fromUtf8(msg).c_str());
    thread->appendLastError(&err);
  }

  // Errors and warnings of one operation (UTF-8). Operations only collect
  // them, so they can run on a worker thread; report() hands them to the
  // script on the CTRL thread.
  struct Messages
  {
    std::vector<std::string> list;

    void add(const std::string &msg) { list.push_back(msg); }

    // Record the exception currently being handled; call only from a catch
    // block. context is UTF-8.
    void addCurrentException(const std::string &context)
    {
      std::string prefix = context.empty() ? std::string() : context + ": ";
      try
      {
        throw;
      }
      catch (const std::exception &e)
      {
        add(prefix + e.what());
      }
      catch (...)
      {
        add(prefix + "unknown error");
      }
    }

    void report(CtrlThread *thread, const char *funcName) const
    {
      for ( const auto &msg : list )
        reportError(thread, funcName, msg);
    }
  };

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

  //----------------------------------------------------------------------------
  // Operations. They take plain inputs (UTF-8 names), write their result into
  // Variables they are given and collect messages, without touching the
  // CtrlThread, so both the blocking functions and the *Async variants (on a
  // worker thread) use them.
  //----------------------------------------------------------------------------

  struct ReadOptions
  {
    bool skipHidden = true;
    bool useHeaders = true;
    bool skipEmpty  = true;
  };

  // Warn about a sheet name the project encoding cannot represent: the
  // returned (converted) name will not find the sheet again.
  void checkRepresentableSheetName(Messages &msgs, const std::string &utf8)
  {
    if ( !ExcelXlsxHelpers::isRepresentable(utf8) )
      msgs.add("sheet name '" + utf8
        + "' cannot be represented in the project encoding; it cannot be opened by the returned name");
  }

  void addWarnings(Messages &msgs, const std::string &context,
                   const std::vector<std::string> &warnings)
  {
    for ( const auto &warning : warnings )
      msgs.add(context + ": " + warning);
  }

  void getSheetNames(const std::string &filename, DynVar &names, Messages &msgs)
  {
    names.reset(TEXT_VAR);
    try
    {
      XLDocument doc;
      doc.open(filename);
      // Worksheets only: chartsheets cannot be read by excelReadSheet.
      for ( const auto &name : doc.workbook().worksheetNames() )
      {
        checkRepresentableSheetName(msgs, name);
        names.append(projectText(name));
      }
      doc.close();
    }
    catch (...)
    {
      msgs.addCurrentException(filename);
    }
  }

  void readSheet(const std::string &filename, std::string sheetName,
                 const ReadOptions &opts, DynVar &rows, Messages &msgs)
  {
    rows.reset(MAPPING_VAR);
    const std::string requestedName = sheetName;
    try
    {
      XLDocument doc;
      doc.open(filename);

      // An empty name means the first worksheet; worksheet(1) would be the
      // first sheet of any kind and fail for a chart sheet.
      if ( sheetName.empty() )
      {
        auto names = doc.workbook().worksheetNames();
        if ( names.empty() )
          throw std::runtime_error("workbook has no worksheets");
        sheetName = names.front();
      }
      auto wks = doc.workbook().worksheet(sheetName);

      std::vector<std::string> warnings;
      ExcelXlsxHelpers::readSheetRows(wks, doc, rows, opts.useHeaders, opts.skipHidden,
                                      opts.skipEmpty, warnings);
      addWarnings(msgs, filename + " [" + sheetName + "]", warnings);
      doc.close();
    }
    catch (...)
    {
      msgs.addCurrentException(filename + " [" + requestedName + "]");
    }
  }

  // Sheets of a file read: key (project encoding) and rows.
  using SheetResults = std::vector<std::pair<std::string, std::unique_ptr<DynVar>>>;

  void readFile(const std::string &filename, const ReadOptions &opts,
                SheetResults &sheets, Messages &msgs)
  {
    sheets.clear();
    try
    {
      XLDocument doc;
      doc.open(filename);

      // Keys in the project encoding; names that collapse to the same text
      // (characters the codepage lacks) get a suffix instead of overwriting
      // each other, as header keys do.
      std::unordered_set<std::string> usedKeys;

      for ( const auto &sn : doc.workbook().worksheetNames() )
      {
        // A broken sheet must not discard the sheets read so far or after it.
        try
        {
          auto wks = doc.workbook().worksheet(sn);

          std::unique_ptr<DynVar> rows(new DynVar);
          rows->reset(MAPPING_VAR);
          std::vector<std::string> warnings;
          ExcelXlsxHelpers::readSheetRows(wks, doc, *rows, opts.useHeaders, opts.skipHidden,
                                          opts.skipEmpty, warnings);
          addWarnings(msgs, filename + " [" + sn + "]", warnings);

          checkRepresentableSheetName(msgs, sn);

          std::string baseKey = ExcelXlsxHelpers::fromUtf8(sn).c_str();
          std::string key = baseKey;
          for ( int n = 2; !usedKeys.insert(key).second; ++n )
            key = baseKey + "_" + std::to_string(n);
          if ( key != baseKey )
            msgs.add("sheet '" + sn + "' is returned under key '"
              + ExcelXlsxHelpers::toUtf8(key.c_str()) + "' because its name collides after encoding conversion");

          sheets.emplace_back(key, std::move(rows));
        }
        catch (...)
        {
          msgs.addCurrentException(filename + " [" + sn + "]");
        }
      }

      doc.close();
    }
    catch (...)
    {
      msgs.addCurrentException(filename);
    }
  }

  // Sheet name (UTF-8) and its rows (a dyn of mappings, possibly
  // anytype-wrapped).
  using SheetList = std::vector<std::pair<std::string, const Variable *>>;

  // Write sheets to a new workbook at filename (UTF-8). The file is only
  // saved when every sheet is valid, so a failed write leaves an existing
  // file untouched.
  bool writeWorkbook(const std::string &filename, const SheetList &sheets, Messages &msgs)
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
      msgs.add(nameProblem);
      return false;
    }

    std::vector<const DynVar *> sheetRows;
    sheetRows.reserve(sheets.size());
    for ( const auto &sheet : sheets )
    {
      const Variable *rows = ExcelXlsxHelpers::unwrapAnyOrMixed(sheet.second);
      if ( !rows || !rows->isDynVar() )
      {
        msgs.add(sheet.first + ": data must be a dyn_mapping");
        return false;
      }
      sheetRows.push_back(static_cast<const DynVar *>(rows));
    }

    try
    {
      if ( !isWritableOrMissing(filename) )
      {
        msgs.add(filename + ": file is locked or not writable");
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
          msgs.add(sheetName + ": " + error);
          return false;
        }
      }

      doc.save();
      doc.close();
      return true;
    }
    catch (...)
    {
      msgs.addCurrentException(filename);
      return false;
    }
  }

  //----------------------------------------------------------------------------
  // Argument parsing, shared by the blocking and the *Async variants. Each
  // returns false after reporting a problem.
  //----------------------------------------------------------------------------

  bool parseReadOptions(ExprList *args, CtrlThread *thread, ReadOptions &opts)
  {
    BitVar skipHidden(true), headers(true), skipEmpty(true);
    if ( !evalOptionalArg(args->getNext(), thread, skipHidden)
      || !evalOptionalArg(args->getNext(), thread, headers)
      || !evalOptionalArg(args->getNext(), thread, skipEmpty) )
      return false;

    opts.skipHidden = skipHidden.isTrue();
    opts.useHeaders = headers.isTrue();
    opts.skipEmpty  = skipEmpty.isTrue();
    return true;
  }

  // Rows argument of excelWriteSheet; owned is set to a copy when the caller
  // needs one (async: the script's variable may change while the worker runs).
  const Variable *evalWriteSheetData(CtrlExpr *expr, CtrlThread *thread)
  {
    return expr ? expr->evaluate(thread) : nullptr;
  }

  // Sheets argument of excelWriteFile, checked instead of copied into a
  // MappingVar: that conversion turns any other type (e.g. a dyn_mapping
  // meant for excelWriteSheet) into an empty mapping, which would "succeed"
  // without writing anything.
  const MappingVar *evalWriteFileData(CtrlExpr *expr, CtrlThread *thread)
  {
    const Variable *value = ExcelXlsxHelpers::unwrapAnyOrMixed(
      expr ? expr->evaluate(thread) : nullptr);
    return (value && value->isA() == MAPPING_VAR)
      ? static_cast<const MappingVar *>(value)
      : nullptr;
  }

  // Sheet list for excelWriteFile. An empty mapping writes a workbook with
  // one empty sheet, as excelWriteSheet does for empty data, instead of no
  // file at all; noRows must outlive the returned list.
  SheetList sheetsOf(const MappingVar &data, const DynVar &noRows)
  {
    SheetList sheets;
    unsigned int numSheets = data.getNumberOfItems();
    if ( numSheets == 0 )
      sheets.emplace_back("Sheet1", &noRows);
    for ( unsigned int s = 0; s < numSheets; s++ )
    {
      CharString sheetName = data.getKey(s)->formatValue(CharString());
      sheets.emplace_back(ExcelXlsxHelpers::toUtf8(sheetName.c_str()), data.getValue(s));
    }
    return sheets;
  }

  //----------------------------------------------------------------------------
  // Async execution: the operation runs on a worker thread while only the
  // calling script waits (WaitCond); the CTRL manager keeps running all other
  // scripts. The result is handed over in checkDone(), on the CTRL thread.
  //----------------------------------------------------------------------------

  struct Job
  {
    std::atomic<bool> done{false};
    Messages msgs;
    std::function<void(Job &)> work;          // runs on the worker thread
    std::function<void(Variable &)> deliver;  // runs on the CTRL thread
  };

  class JobWait : public WaitCond
  {
    public:
      JobWait(std::shared_ptr<Job> job, CtrlThread *thread, Variable *target, std::string funcName)
        : job_(std::move(job)), thread_(thread), target_(target), funcName_(std::move(funcName)) { }

      // Poll interval: checkDone() is cheap, but the engine need not call it
      // more often than this.
      const TimeVar &nextCheck() const override
      {
        next_ = TimeVar();
        next_ += TimeVar(0, static_cast<PVSSshort>(POLL_MILLIS));
        return next_;
      }

      int checkDone() override
      {
        if ( !job_->done.load(std::memory_order_acquire) )
          return 0;

        if ( !delivered_ )
        {
          delivered_ = true;
          job_->deliver(*target_);
          job_->msgs.report(thread_, funcName_.c_str());
        }
        return 1;
      }

    private:
      static constexpr int POLL_MILLIS = 50;

      // Shared with the worker thread, which may outlive this wait condition
      // (e.g. when the script is stopped while waiting).
      std::shared_ptr<Job> job_;
      CtrlThread *thread_;
      Variable *target_;
      std::string funcName_; // copied: param.funcName does not outlive execute()
      bool delivered_ = false;
      mutable TimeVar next_{0, static_cast<PVSSshort>(0)};
  };

  // Hand rows (a dyn of mappings) to the script's variable. Moved when the
  // target is a dyn_mapping, so a huge result is not deep-copied on the CTRL
  // thread; any other type gets CTRL's converting assignment.
  void deliverRows(DynVar &rows, Variable &target)
  {
    if ( target.isA() == DYNMAPPING_VAR )
      rows.moveAllItems(static_cast<DynVar &>(target));
    else
      target = rows;
  }

  void startJob(CtrlThread *thread, Variable *target, const char *funcName,
                std::shared_ptr<Job> job)
  {
    thread->setWaitCond(new JobWait(job, thread, target, funcName));

    // Detached: the job state is shared, so it stays valid if the script
    // (and its JobWait) goes away first.
    std::thread([job]
    {
      try
      {
        job->work(*job);
      }
      catch (...)
      {
        job->msgs.addCurrentException(std::string());
      }
      job->done.store(true, std::memory_order_release);
    }).detach();
  }
}

//------------------------------------------------------------------------------

const Variable *ExternHdl::execute(ExecuteParamRec &param)
{
  enum
  {
    F_excelGetSheetNames      = 0,
    F_excelReadSheet          = 1,
    F_excelReadFile           = 2,
    F_excelWriteSheet         = 3,
    F_excelWriteFile          = 4,
    F_excelGetSheetNamesAsync = 5,
    F_excelReadSheetAsync     = 6,
    F_excelReadFileAsync      = 7,
    F_excelWriteSheetAsync    = 8,
    F_excelWriteFileAsync     = 9,
  };

  static DynVar dynTextResult;
  static DynVar dynMappingResult;
  static MappingVar mappingResult;
  static BitVar writeResult;

  // *Async results: 0 = started (outcome via the reference parameter and
  // getLastError() once the call returns), -1 = invalid arguments.
  static IntegerVar asyncStarted(0);
  static IntegerVar asyncRejected(-1);

  const char *funcName = param.funcName.c_str();
  param.thread->clearLastError();

  switch ( param.funcNum )
  {
    // -------------------------------------------------------------------------
    // excelGetSheetNames(string filename) -> dyn_string
    case F_excelGetSheetNames:
    {
      dynTextResult.reset(TEXT_VAR);

      TextVar filenameVar;
      if ( !hasNumArgs(1, 1, param) )
        return &dynTextResult;
      if ( !evalArg(param.args->getFirst(), param.thread, filenameVar) )
      {
        reportError(param.thread, funcName, "argument could not be evaluated");
        return &dynTextResult;
      }

      Messages msgs;
      getSheetNames(utf8Arg(filenameVar), dynTextResult, msgs);
      msgs.report(param.thread, funcName);
      return &dynTextResult;
    }

    // -------------------------------------------------------------------------
    // excelReadSheet(string filename, string sheetName,
    //                bool skipHiddenRows, bool firstRowIsColumnNames,
    //                bool skipEmptyRows) -> dyn_mapping
    case F_excelReadSheet:
    {
      dynMappingResult.reset(MAPPING_VAR);

      TextVar filenameVar, sheetnameVar;
      ReadOptions opts;
      if ( !hasNumArgs(2, 5, param) )
        return &dynMappingResult;
      if ( !evalArg(param.args->getFirst(), param.thread, filenameVar)
        || !evalArg(param.args->getNext(), param.thread, sheetnameVar)
        || !parseReadOptions(param.args, param.thread, opts) )
      {
        reportError(param.thread, funcName, "argument could not be evaluated");
        return &dynMappingResult;
      }

      Messages msgs;
      readSheet(utf8Arg(filenameVar), utf8Arg(sheetnameVar), opts, dynMappingResult, msgs);
      msgs.report(param.thread, funcName);
      return &dynMappingResult;
    }

    // -------------------------------------------------------------------------
    // excelReadFile(string filename, bool skipHiddenRows,
    //              bool firstRowIsColumnNames, bool skipEmptyRows) -> mapping
    case F_excelReadFile:
    {
      mappingResult = MappingVar();

      TextVar filenameVar;
      ReadOptions opts;
      if ( !hasNumArgs(1, 4, param) )
        return &mappingResult;
      if ( !evalArg(param.args->getFirst(), param.thread, filenameVar)
        || !parseReadOptions(param.args, param.thread, opts) )
      {
        reportError(param.thread, funcName, "argument could not be evaluated");
        return &mappingResult;
      }

      Messages msgs;
      SheetResults sheets;
      readFile(utf8Arg(filenameVar), opts, sheets, msgs);
      for ( auto &sheet : sheets )
        mappingResult.setAt(TextVar(sheet.first.c_str()), *sheet.second);
      msgs.report(param.thread, funcName);
      return &mappingResult;
    }

    // -------------------------------------------------------------------------
    // excelWriteSheet(string filename, string sheetName, dyn_anytype data)
    //   -> bool
    case F_excelWriteSheet:
    {
      writeResult = BitVar(false);

      TextVar filenameVar, sheetnameVar;
      if ( !hasNumArgs(3, 3, param) )
        return &writeResult;
      if ( !evalArg(param.args->getFirst(), param.thread, filenameVar)
        || !evalArg(param.args->getNext(), param.thread, sheetnameVar) )
      {
        reportError(param.thread, funcName, "argument could not be evaluated");
        return &writeResult;
      }

      // Evaluated last and not copied; writeWorkbook reports nullptr or a
      // non-dyn value as invalid data.
      const Variable *data = evalWriteSheetData(param.args->getNext(), param.thread);
      std::string sheetname = utf8Arg(sheetnameVar);

      Messages msgs;
      writeResult = BitVar(writeWorkbook(utf8Arg(filenameVar),
        SheetList{ { sheetname.empty() ? "Sheet1" : sheetname, data } }, msgs));
      msgs.report(param.thread, funcName);
      return &writeResult;
    }

    // -------------------------------------------------------------------------
    // excelWriteFile(string filename, mapping data) -> bool
    case F_excelWriteFile:
    {
      writeResult = BitVar(false);

      TextVar filenameVar;
      if ( !hasNumArgs(2, 2, param) )
        return &writeResult;
      if ( !evalArg(param.args->getFirst(), param.thread, filenameVar) )
      {
        reportError(param.thread, funcName, "argument could not be evaluated");
        return &writeResult;
      }

      const MappingVar *data = evalWriteFileData(param.args->getNext(), param.thread);
      if ( !data )
      {
        reportError(param.thread, funcName, "data must be a mapping of sheet name to dyn_mapping");
        return &writeResult;
      }

      DynVar noRows;
      noRows.reset(MAPPING_VAR);

      Messages msgs;
      writeResult = BitVar(writeWorkbook(utf8Arg(filenameVar), sheetsOf(*data, noRows), msgs));
      msgs.report(param.thread, funcName);
      return &writeResult;
    }

    // -------------------------------------------------------------------------
    // excelGetSheetNamesAsync(string filename, dyn_string &names) -> int
    case F_excelGetSheetNamesAsync:
    {
      TextVar filenameVar;
      if ( !hasNumArgs(2, 2, param) )
        return &asyncRejected;
      Variable *target = nullptr;
      if ( !evalArg(param.args->getFirst(), param.thread, filenameVar)
        || !(target = getTarget(param.args->getNext(), param, NO_VAR)) )
      {
        reportError(param.thread, funcName, "argument could not be evaluated or is not a variable");
        return &asyncRejected;
      }

      auto job = std::make_shared<Job>();
      auto names = std::make_shared<DynVar>();
      std::string filename = utf8Arg(filenameVar);
      job->work = [filename, names](Job &j) { getSheetNames(filename, *names, j.msgs); };
      job->deliver = [names](Variable &out) { out = *names; };
      startJob(param.thread, target, funcName, job);
      return &asyncStarted;
    }

    // -------------------------------------------------------------------------
    // excelReadSheetAsync(string filename, string sheetName, dyn_mapping &rows,
    //                     bool skipHiddenRows, bool firstRowIsColumnNames,
    //                     bool skipEmptyRows) -> int
    case F_excelReadSheetAsync:
    {
      TextVar filenameVar, sheetnameVar;
      ReadOptions opts;
      if ( !hasNumArgs(3, 6, param) )
        return &asyncRejected;
      Variable *target = nullptr;
      if ( !evalArg(param.args->getFirst(), param.thread, filenameVar)
        || !evalArg(param.args->getNext(), param.thread, sheetnameVar)
        || !(target = getTarget(param.args->getNext(), param, NO_VAR))
        || !parseReadOptions(param.args, param.thread, opts) )
      {
        reportError(param.thread, funcName, "argument could not be evaluated or is not a variable");
        return &asyncRejected;
      }

      auto job = std::make_shared<Job>();
      auto rows = std::make_shared<DynVar>();
      std::string filename = utf8Arg(filenameVar);
      std::string sheetname = utf8Arg(sheetnameVar);
      job->work = [filename, sheetname, opts, rows](Job &j)
      {
        readSheet(filename, sheetname, opts, *rows, j.msgs);
      };
      job->deliver = [rows](Variable &out) { deliverRows(*rows, out); };
      startJob(param.thread, target, funcName, job);
      return &asyncStarted;
    }

    // -------------------------------------------------------------------------
    // excelReadFileAsync(string filename, mapping &sheets, bool skipHiddenRows,
    //                    bool firstRowIsColumnNames, bool skipEmptyRows) -> int
    case F_excelReadFileAsync:
    {
      TextVar filenameVar;
      ReadOptions opts;
      if ( !hasNumArgs(2, 5, param) )
        return &asyncRejected;
      Variable *target = nullptr;
      if ( !evalArg(param.args->getFirst(), param.thread, filenameVar)
        || !(target = getTarget(param.args->getNext(), param, NO_VAR))
        || !parseReadOptions(param.args, param.thread, opts) )
      {
        reportError(param.thread, funcName, "argument could not be evaluated or is not a variable");
        return &asyncRejected;
      }

      auto job = std::make_shared<Job>();
      auto sheets = std::make_shared<SheetResults>();
      std::string filename = utf8Arg(filenameVar);
      job->work = [filename, opts, sheets](Job &j) { readFile(filename, opts, *sheets, j.msgs); };
      job->deliver = [sheets](Variable &out)
      {
        if ( out.isA() != MAPPING_VAR )
        {
          // Unusual target type: let CTRL's assignment convert a copy.
          MappingVar result;
          for ( auto &sheet : *sheets )
            result.setAt(TextVar(sheet.first.c_str()), *sheet.second);
          out = result;
          return;
        }

        MappingVar &result = static_cast<MappingVar &>(out);
        result = MappingVar();
        for ( auto &sheet : *sheets )
        {
          // Insert an empty dyn_mapping, then move the rows into it.
          TextVar key(sheet.first.c_str());
          DynVar empty;
          empty.reset(MAPPING_VAR);
          result.setAt(key, empty);
          Variable *slot = result.getAt(key);
          deliverRows(*sheet.second, slot ? *slot : static_cast<Variable &>(empty));
        }
      };
      startJob(param.thread, target, funcName, job);
      return &asyncStarted;
    }

    // -------------------------------------------------------------------------
    // excelWriteSheetAsync(string filename, string sheetName, dyn_anytype data,
    //                      bool &ok) -> int
    case F_excelWriteSheetAsync:
    {
      TextVar filenameVar, sheetnameVar;
      if ( !hasNumArgs(4, 4, param) )
        return &asyncRejected;
      if ( !evalArg(param.args->getFirst(), param.thread, filenameVar)
        || !evalArg(param.args->getNext(), param.thread, sheetnameVar) )
      {
        reportError(param.thread, funcName, "argument could not be evaluated");
        return &asyncRejected;
      }

      // Copied: the script's variable may change while the worker runs.
      const Variable *data = evalWriteSheetData(param.args->getNext(), param.thread);
      std::shared_ptr<Variable> dataCopy(data ? data->clone() : nullptr);

      Variable *target = getTarget(param.args->getNext(), param, NO_VAR);
      if ( !target )
      {
        reportError(param.thread, funcName, "ok must be a variable");
        return &asyncRejected;
      }

      auto job = std::make_shared<Job>();
      auto ok = std::make_shared<bool>(false);
      std::string filename = utf8Arg(filenameVar);
      std::string sheetname = utf8Arg(sheetnameVar);
      if ( sheetname.empty() )
        sheetname = "Sheet1";
      job->work = [filename, sheetname, dataCopy, ok](Job &j)
      {
        *ok = writeWorkbook(filename, SheetList{ { sheetname, dataCopy.get() } }, j.msgs);
      };
      job->deliver = [ok](Variable &out) { out = BitVar(*ok); };
      startJob(param.thread, target, funcName, job);
      return &asyncStarted;
    }

    // -------------------------------------------------------------------------
    // excelWriteFileAsync(string filename, mapping data, bool &ok) -> int
    case F_excelWriteFileAsync:
    {
      TextVar filenameVar;
      if ( !hasNumArgs(3, 3, param) )
        return &asyncRejected;
      if ( !evalArg(param.args->getFirst(), param.thread, filenameVar) )
      {
        reportError(param.thread, funcName, "argument could not be evaluated");
        return &asyncRejected;
      }

      const MappingVar *data = evalWriteFileData(param.args->getNext(), param.thread);
      if ( !data )
      {
        reportError(param.thread, funcName, "data must be a mapping of sheet name to dyn_mapping");
        return &asyncRejected;
      }
      // Copied: the script's variable may change while the worker runs.
      std::shared_ptr<Variable> dataCopy(data->clone());

      Variable *target = getTarget(param.args->getNext(), param, NO_VAR);
      if ( !target )
      {
        reportError(param.thread, funcName, "ok must be a variable");
        return &asyncRejected;
      }

      auto job = std::make_shared<Job>();
      auto ok = std::make_shared<bool>(false);
      std::string filename = utf8Arg(filenameVar);
      job->work = [filename, dataCopy, ok](Job &j)
      {
        DynVar noRows;
        noRows.reset(MAPPING_VAR);
        *ok = writeWorkbook(filename,
          sheetsOf(*static_cast<const MappingVar *>(dataCopy.get()), noRows), j.msgs);
      };
      job->deliver = [ok](Variable &out) { out = BitVar(*ok); };
      startJob(param.thread, target, funcName, job);
      return &asyncStarted;
    }

    // -------------------------------------------------------------------------
    default:
      return &errorIntVar;
  }
}

//------------------------------------------------------------------------------
