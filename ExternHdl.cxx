#include <ExternHdl.hxx>

#include <ExcelXlsxHelpers.hxx>

#include <BitVar.hxx>
#include <DynVar.hxx>
#include <ErrClass.hxx>
#include <IntegerVar.hxx>
#include <MappingVar.hxx>
#include <TimeVar.hxx>
#include <WaitCond.hxx>

#include <algorithm>
#include <atomic>
#include <cctype>
#include <chrono>
#include <condition_variable>
#include <deque>
#include <exception>
#include <filesystem>
#include <fstream>
#include <functional>
#include <memory>
#include <mutex>
#include <shared_mutex>
#include <stdexcept>
#include <string>
#include <system_error>
#include <thread>
#include <unordered_map>
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

  // Serialises operations on the same file: writes exclusive, reads shared.
  // The *Async variants run on worker threads, so two scripts writing (or
  // one writing, one reading) the same path would otherwise race on save();
  // the single-threaded CTRL engine used to serialise them implicitly.
  class PathLock
  {
    public:
      PathLock(const std::string &utf8Path, bool exclusive)
        : mutex_(mutexFor(utf8Path)), exclusive_(exclusive)
      {
        if ( exclusive_ )
          mutex_->lock();
        else
          mutex_->lock_shared();
      }

      ~PathLock()
      {
        if ( exclusive_ )
          mutex_->unlock();
        else
          mutex_->unlock_shared();
      }

      PathLock(const PathLock &) = delete;
      PathLock &operator=(const PathLock &) = delete;

    private:
      static std::string normalisedPath(const std::string &utf8Path)
      {
        std::string key;
        try
        {
          std::error_code ec;
          std::filesystem::path path = std::filesystem::u8path(utf8Path);
          std::filesystem::path full = std::filesystem::weakly_canonical(path, ec);
          key = (ec ? path : full).lexically_normal().u8string();
        }
        catch (const std::exception &)
        {
          key = utf8Path; // e.g. invalid UTF-8: lock on the literal name
        }
#ifdef _WIN32
        // Windows paths are case-insensitive.
        for ( char &c : key )
          c = static_cast<char>(tolower(static_cast<unsigned char>(c)));
#endif
        return key;
      }

      static std::shared_ptr<std::shared_mutex> mutexFor(const std::string &utf8Path)
      {
        // Never destroyed: worker threads may still take locks while static
        // objects are torn down at manager exit.
        static std::mutex *registryMutex = new std::mutex;
        static auto *registry = new std::unordered_map<std::string, std::weak_ptr<std::shared_mutex>>;

        std::string key = normalisedPath(utf8Path);
        std::lock_guard<std::mutex> guard(*registryMutex);

        for ( auto it = registry->begin(); it != registry->end(); )
          it = it->second.expired() ? registry->erase(it) : std::next(it);

        std::shared_ptr<std::shared_mutex> mutex = (*registry)[key].lock();
        if ( !mutex )
        {
          mutex = std::make_shared<std::shared_mutex>();
          (*registry)[key] = mutex;
        }
        return mutex;
      }

      std::shared_ptr<std::shared_mutex> mutex_;
      bool exclusive_;
  };

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
      PathLock lock(filename, false);
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
      PathLock lock(filename, false);
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
      PathLock lock(filename, false);
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
      PathLock lock(filename, true);
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
    std::function<void(Job &)> work;          // runs on a worker thread
    std::function<void(Variable &)> deliver;  // runs on the CTRL thread
  };

  // A few worker threads shared by all *Async calls, started on demand.
  // Bounds concurrency (each running job holds a whole workbook in memory)
  // and thread creation (a failure is reported instead of leaving a script
  // waiting). At manager exit queued jobs are dropped and the workers are
  // joined, so none runs while the extension is torn down.
  class WorkerPool
  {
    public:
      static WorkerPool &instance()
      {
        static WorkerPool pool;
        return pool;
      }

      // Queue a job; false if no worker thread exists or could be started.
      bool submit(std::shared_ptr<Job> job)
      {
        std::unique_lock<std::mutex> lock(mutex_);
        if ( stopping_ )
          return false;

        if ( idle_ == 0 && threads_.size() < maxThreads() )
        {
          try
          {
            threads_.emplace_back([this] { run(); });
          }
          catch (const std::system_error &)
          {
            if ( threads_.empty() )
              return false; // no thread at all; busy workers would pick it up otherwise
          }
        }

        queue_.push_back(std::move(job));
        lock.unlock();
        wakeup_.notify_one();
        return true;
      }

      ~WorkerPool()
      {
        {
          std::lock_guard<std::mutex> lock(mutex_);
          stopping_ = true;
          queue_.clear();
        }
        wakeup_.notify_all();
        for ( auto &thread : threads_ )
          if ( thread.joinable() )
            thread.join();
      }

    private:
      WorkerPool() = default;

      static size_t maxThreads()
      {
        unsigned cores = std::thread::hardware_concurrency();
        return std::clamp<size_t>(cores ? cores / 2 : 2, 2, 4);
      }

      void run()
      {
        for (;;)
        {
          std::shared_ptr<Job> job;
          {
            std::unique_lock<std::mutex> lock(mutex_);
            ++idle_;
            wakeup_.wait(lock, [this] { return stopping_ || !queue_.empty(); });
            --idle_;
            if ( stopping_ )
              return;
            job = std::move(queue_.front());
            queue_.pop_front();
          }

          try
          {
            job->work(*job);
          }
          catch (...)
          {
            job->msgs.addCurrentException(std::string());
          }
          job->done.store(true, std::memory_order_release);
        }
      }

      std::mutex mutex_;
      std::condition_variable wakeup_;
      std::deque<std::shared_ptr<Job>> queue_;
      std::vector<std::thread> threads_;
      size_t idle_ = 0;
      bool stopping_ = false;
  };

  class JobWait : public WaitCond
  {
    public:
      JobWait(std::shared_ptr<Job> job, const ExternHdl &hdl,
              const BaseExternHdl::ExecuteParamRec &param, CtrlExpr *targetExpr)
        : job_(std::move(job)), hdl_(hdl), param_(param), targetExpr_(targetExpr),
          funcName_(param.funcName.c_str()), started_(std::chrono::steady_clock::now()) { }

      // Poll interval: short at first so small files return quickly, then
      // relaxed for long jobs. The worker cannot wake the engine.
      const TimeVar &nextCheck() const override
      {
        auto elapsed = std::chrono::steady_clock::now() - started_;
        PVSSshort interval = elapsed < std::chrono::milliseconds(100) ? 5 : 50;
        next_ = TimeVar();
        next_ += TimeVar(0, interval);
        return next_;
      }

      int checkDone() override
      {
        if ( !job_->done.load(std::memory_order_acquire) )
          return 0;

        if ( !delivered_ )
        {
          delivered_ = true;

          // Nothing may escape into the engine's scheduler.
          try
          {
            // Resolved again instead of keeping the pointer from the call:
            // the variable (e.g. an element of a shared dyn) may have been
            // removed by another script while this one waited.
            Variable *target = hdl_.resolveTarget(targetExpr_, param_);
            if ( target )
              job_->deliver(*target);
            else
              job_->msgs.add("the result variable no longer exists; result discarded");
          }
          catch (...)
          {
            job_->msgs.addCurrentException("delivering the result");
          }

          try
          {
            job_->msgs.report(param_.thread, funcName_.c_str());
          }
          catch (...)
          {
          }
        }
        return 1;
      }

    private:
      // Shared with the worker thread, which may still hold it if the script
      // (and this wait condition) goes away first.
      std::shared_ptr<Job> job_;
      const ExternHdl &hdl_;
      BaseExternHdl::ExecuteParamRec param_;
      CtrlExpr *targetExpr_;
      std::string funcName_; // copied: param.funcName does not outlive execute()
      std::chrono::steady_clock::time_point started_;
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

  // Hand the sheets of readFile() to a mapping variable, moving the rows.
  void deliverSheets(SheetResults &sheets, Variable &target)
  {
    if ( target.isA() != MAPPING_VAR )
    {
      // Unusual target type: let CTRL's assignment convert a copy.
      MappingVar result;
      for ( auto &sheet : sheets )
        result.setAt(TextVar(sheet.first.c_str()), *sheet.second);
      target = result;
      return;
    }

    MappingVar &result = static_cast<MappingVar &>(target);
    result = MappingVar();
    for ( auto &sheet : sheets )
    {
      // Insert an empty dyn_mapping, then move the rows into it.
      TextVar key(sheet.first.c_str());
      DynVar empty;
      empty.reset(MAPPING_VAR);
      result.setAt(key, empty);
      Variable *slot = result.getAt(key);
      deliverRows(*sheet.second, slot ? *slot : static_cast<Variable &>(empty));
    }
  }

  // Queue the job and suspend the calling script until it is done. False
  // (nothing queued, script not suspended) if no worker thread can run it.
  bool startJob(const ExternHdl &hdl, const BaseExternHdl::ExecuteParamRec &param,
                CtrlExpr *targetExpr, std::shared_ptr<Job> job)
  {
    if ( !WorkerPool::instance().submit(job) )
      return false;

    param.thread->setWaitCond(new JobWait(job, hdl, param, targetExpr));
    return true;
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
      deliverSheets(sheets, mappingResult);
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
      CtrlExpr *targetExpr = nullptr;
      if ( !evalArg(param.args->getFirst(), param.thread, filenameVar)
        || !resolveTarget(targetExpr = param.args->getNext(), param) )
      {
        reportError(param.thread, funcName, "argument could not be evaluated or is not a variable");
        return &asyncRejected;
      }

      auto job = std::make_shared<Job>();
      auto names = std::make_shared<DynVar>();
      std::string filename = utf8Arg(filenameVar);
      job->work = [filename, names](Job &j) { getSheetNames(filename, *names, j.msgs); };
      job->deliver = [names](Variable &out) { out = *names; };
      if ( !startJob(*this, param, targetExpr, job) )
      {
        reportError(param.thread, funcName, "no worker thread available");
        return &asyncRejected;
      }
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
      CtrlExpr *targetExpr = nullptr;
      if ( !evalArg(param.args->getFirst(), param.thread, filenameVar)
        || !evalArg(param.args->getNext(), param.thread, sheetnameVar)
        || !resolveTarget(targetExpr = param.args->getNext(), param)
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
      if ( !startJob(*this, param, targetExpr, job) )
      {
        reportError(param.thread, funcName, "no worker thread available");
        return &asyncRejected;
      }
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
      CtrlExpr *targetExpr = nullptr;
      if ( !evalArg(param.args->getFirst(), param.thread, filenameVar)
        || !resolveTarget(targetExpr = param.args->getNext(), param)
        || !parseReadOptions(param.args, param.thread, opts) )
      {
        reportError(param.thread, funcName, "argument could not be evaluated or is not a variable");
        return &asyncRejected;
      }

      auto job = std::make_shared<Job>();
      auto sheets = std::make_shared<SheetResults>();
      std::string filename = utf8Arg(filenameVar);
      job->work = [filename, opts, sheets](Job &j) { readFile(filename, opts, *sheets, j.msgs); };
      job->deliver = [sheets](Variable &out) { deliverSheets(*sheets, out); };
      if ( !startJob(*this, param, targetExpr, job) )
      {
        reportError(param.thread, funcName, "no worker thread available");
        return &asyncRejected;
      }
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

      CtrlExpr *targetExpr = param.args->getNext();
      if ( !resolveTarget(targetExpr, param) )
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
      if ( !startJob(*this, param, targetExpr, job) )
      {
        reportError(param.thread, funcName, "no worker thread available");
        return &asyncRejected;
      }
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

      CtrlExpr *targetExpr = param.args->getNext();
      if ( !resolveTarget(targetExpr, param) )
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
      if ( !startJob(*this, param, targetExpr, job) )
      {
        reportError(param.thread, funcName, "no worker thread available");
        return &asyncRejected;
      }
      return &asyncStarted;
    }

    // -------------------------------------------------------------------------
    default:
      return &errorIntVar;
  }
}

//------------------------------------------------------------------------------
