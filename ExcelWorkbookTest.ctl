// $License: NOLICENSE
//--------------------------------------------------------------------------------
/**
  @file $relPath
  @copyright $copyright
  @author Kilian von Pflugk

  Tests of the CTRL++ classes in scripts/libs/classes/ExcelWorkbook.ctl.
  Needs the CtrlExcelReader extension and the library in the project; copy
  testdata/CtrlExcelReaderFixture.xlsx into the project's data directory for
  excelWorkbookFixtureTest (skipped otherwise).
*/

//--------------------------------------------------------------------------------
// Libraries used (#uses)
#uses "classes/ExcelWorkbook"

//--------------------------------------------------------------------------------
// Variables and Constants

// Results collected by recordTest() / skipTest() for the summary in main().
int g_passedTests;
dyn_string g_failedTests;
dyn_string g_skippedTests;

//--------------------------------------------------------------------------------
void main()
{
  recordTest("excelSheetRowsTest", excelSheetRowsTest());
  recordTest("excelSheetColumnsTest", excelSheetColumnsTest());
  recordTest("excelWorkbookSheetsTest", excelWorkbookSheetsTest());
  recordTest("excelWorkbookRoundTripAsyncTest", excelWorkbookRoundTripTest(ExcelCallMode::Async,
                                                                           "excelWorkbookRoundTripAsyncTest"));
  recordTest("excelWorkbookRoundTripBlockingTest", excelWorkbookRoundTripTest(ExcelCallMode::Blocking,
                                                                              "excelWorkbookRoundTripBlockingTest"));
  recordTest("excelWorkbookLoadSheetTest", excelWorkbookLoadSheetTest());
  recordTest("excelWorkbookReadOptionsTest", excelWorkbookReadOptionsTest());
  recordTest("excelWorkbookErrorsTest", excelWorkbookErrorsTest());
  recordTest("excelWorkbookFixtureTest", excelWorkbookFixtureTest());

  DebugTN("ExcelWorkbookTest summary",
          dynlen(g_failedTests) == 0 ? "ALL PASSED" : "FAILURES",
          "passed", g_passedTests,
          "failed", g_failedTests,
          "skipped", g_skippedTests);
}

// Counts a test result; tests that called skipTest() count as skipped.
void recordTest(string name, bool pass)
{
  if (dynContains(g_skippedTests, name) > 0)
    return;

  if (pass)
    g_passedTests++;
  else
    dynAppend(g_failedTests, name);
}

// Marks a test as skipped (it then returns TRUE, which recordTest ignores).
void skipTest(string name, string reason)
{
  dynAppend(g_skippedTests, name);
  DebugTN(name, "skipped: " + reason);
}

// Logs the failed checks of a test and returns whether there were none.
bool report(string name, const dyn_string &failed)
{
  bool pass = dynlen(failed) == 0;

  if (pass)
    DebugTN(name, "pass", pass);
  else
    DebugTN(name, "pass", pass, "failed checks", failed);

  return pass;
}

//--------------------------------------------------------------------------------
// ExcelSheet in memory: rows and cells.
bool excelSheetRowsTest()
{
  dyn_string failed;
  shared_ptr<ExcelSheet> sheet = new ExcelSheet("People");

  if (sheet.getName() != "People" || sheet.rowCount() != 0) dynAppend(failed, "new sheet");

  sheet.appendRow(makeMapping("Name", "Alice", "Age", 30));
  sheet.appendRow(makeMapping("Name", "Bob"));

  if (sheet.rowCount() != 2) dynAppend(failed, "rowCount");
  if (sheet.getValue(1, "Age") != 30) dynAppend(failed, "getValue");
  if (sheet.getValue(2, "Age") != "") dynAppend(failed, "getValue of a missing column");
  if (sheet.getValue(3, "Name") != "") dynAppend(failed, "getValue of a missing row");
  if (mappinglen(sheet.getRow(0)) != 0 || mappinglen(sheet.getRow(3)) != 0) dynAppend(failed, "getRow out of range");

  // Setting a cell of row rowCount() + 1 appends a row; beyond that fails.
  if (!sheet.setValue(2, "Age", 25) || sheet.getValue(2, "Age") != 25) dynAppend(failed, "setValue");
  if (!sheet.setValue(3, "Name", "Carol") || sheet.rowCount() != 3) dynAppend(failed, "setValue appends");
  if (sheet.setValue(5, "Name", "Dave") || sheet.rowCount() != 3) dynAppend(failed, "setValue past the end");

  if (!sheet.setRow(1, makeMapping("Name", "Alice", "Age", 31)) || sheet.getValue(1, "Age") != 31)
    dynAppend(failed, "setRow");
  if (!sheet.setRow(4, makeMapping("Name", "Dave")) || sheet.rowCount() != 4) dynAppend(failed, "setRow appends");
  if (sheet.setRow(0, makeMapping()) || sheet.setRow(6, makeMapping())) dynAppend(failed, "setRow out of range");

  if (!sheet.removeRow(4) || sheet.rowCount() != 3) dynAppend(failed, "removeRow");
  if (sheet.removeRow(4)) dynAppend(failed, "removeRow out of range");

  dyn_mapping rows = sheet.getRows();
  if (dynlen(rows) != 3 || rows[3]["Name"] != "Carol") dynAppend(failed, "getRows");

  sheet.clear();
  if (sheet.rowCount() != 0) dynAppend(failed, "clear");

  sheet.setRows(rows);
  sheet.setName("Renamed");
  if (sheet.rowCount() != 3 || sheet.getName() != "Renamed") dynAppend(failed, "setRows / setName");

  return report("excelSheetRowsTest", failed);
}

// ExcelSheet in memory: column access.
bool excelSheetColumnsTest()
{
  dyn_string failed;
  dyn_mapping rows;
  dynAppend(rows, makeMapping("Name", "Alice", "Unit", "A"));
  dynAppend(rows, makeMapping("Name", "Bob", "Unit", "B", "Extra", 1));
  dynAppend(rows, makeMapping("Name", "Carol", "Unit", "A"));
  shared_ptr<ExcelSheet> sheet = new ExcelSheet("Staff", rows);

  dyn_anytype names = sheet.getColumnNames();
  // Within a row the mapping decides the order; "Extra" is first seen in row 2.
  if (dynlen(names) != 3 || dynContains(names, "Name") <= 0 || dynContains(names, "Unit") <= 0 || names[3] != "Extra")
  {
    DebugTN("excelSheetColumnsTest: getColumnNames", names);
    dynAppend(failed, "getColumnNames");
  }

  dyn_anytype extra = sheet.getColumn("Extra");
  if (dynlen(extra) != 3 || extra[1] != "" || extra[2] != 1 || extra[3] != "") dynAppend(failed, "getColumn");

  dyn_mapping unitA = sheet.findRows("Unit", "A");
  if (dynlen(unitA) != 2 || unitA[1]["Name"] != "Alice" || unitA[2]["Name"] != "Carol") dynAppend(failed, "findRows");
  if (dynlen(sheet.findRows("Unit", "C")) != 0 || dynlen(sheet.findRows("Missing", "A")) != 0)
    dynAppend(failed, "findRows without match");

  return report("excelSheetColumnsTest", failed);
}

// ExcelWorkbook: sheets in memory.
bool excelWorkbookSheetsTest()
{
  dyn_string failed;
  shared_ptr<ExcelWorkbook> book = new ExcelWorkbook("unused.xlsx");

  if (book.sheetCount() != 0 || book.getFilename() != "unused.xlsx") dynAppend(failed, "new workbook");
  if (book.getCallMode() != ExcelCallMode::Async) dynAppend(failed, "default call mode");

  shared_ptr<ExcelSheet> a = book.addSheet("A");
  book.addSheet("B");
  a.appendRow(makeMapping("x", 1));

  // addSheet returns the existing sheet; sheets are shared, not copied.
  if (book.addSheet("A").rowCount() != 1 || book.sheetCount() != 2) dynAppend(failed, "addSheet existing");
  if (book.getSheet("A").getValue(1, "x") != 1) dynAppend(failed, "getSheet shares the sheet");
  if (book.getSheet("C") != nullptr) dynAppend(failed, "getSheet missing");
  if (!book.hasSheet("B") || book.hasSheet("b")) dynAppend(failed, "hasSheet");
  if (book.getSheetNames() != makeDynString("A", "B")) dynAppend(failed, "getSheetNames");

  if (!book.removeSheet("A") || book.removeSheet("A") || book.getSheetNames() != makeDynString("B"))
    dynAppend(failed, "removeSheet");

  return report("excelWorkbookSheetsTest", failed);
}

// save() and load() in both call modes; load() keeps the file's sheet order.
bool excelWorkbookRoundTripTest(ExcelCallMode mode, string name)
{
  dyn_string failed;
  string filename = tmpnam();
  time stamp = makeTime(2026, 3, 15, 8, 30, 15, 250);

  shared_ptr<ExcelWorkbook> out = new ExcelWorkbook(filename);
  out.setCallMode(mode);
  shared_ptr<ExcelSheet> people = out.addSheet("People");
  people.appendRow(makeMapping("Name", "Alice", "Age", 30, "Seen", stamp));
  people.appendRow(makeMapping("Name", "Bob", "Age", 25, "Seen", stamp, "Active", TRUE));
  shared_ptr<ExcelSheet> data = out.addSheet("Data");
  data.appendRow(makeMapping("Value", 1.5));

  if (!out.save()) dynAppend(failed, "save: " + out.getMessageText());
  if (out.hasMessages()) dynAppend(failed, "save messages: " + out.getMessageText());

  shared_ptr<ExcelWorkbook> in = new ExcelWorkbook(filename);
  in.setCallMode(mode);

  if (!in.exists()) dynAppend(failed, "exists");
  // excelWriteFile writes the sheets in mapping key order; load() keeps the file's.
  dyn_string fileOrder = in.readSheetNames();
  if (dynlen(fileOrder) != 2 || dynContains(fileOrder, "Data") <= 0 || dynContains(fileOrder, "People") <= 0)
    dynAppend(failed, "readSheetNames");
  if (!in.load()) dynAppend(failed, "load: " + in.getMessageText());
  if (in.hasMessages()) dynAppend(failed, "load messages: " + in.getMessageText());
  if (in.getSheetNames() != fileOrder) dynAppend(failed, "load keeps the file order");

  shared_ptr<ExcelSheet> back = in.getSheet("People");

  if (back == nullptr)
  {
    dynAppend(failed, "People missing");
  }
  else
  {
    if (back.rowCount() != 2) dynAppend(failed, "row count " + back.rowCount());
    if (back.getValue(1, "Name") != "Alice" || back.getValue(2, "Age") != 25) dynAppend(failed, "values");
    if (back.getValue(1, "Seen") != stamp) dynAppend(failed, "time");
    if (back.getValue(2, "Active") != TRUE || back.getValue(1, "Active") != "") dynAppend(failed, "union of keys");
  }

  if (in.getSheet("Data") == nullptr || in.getSheet("Data").getValue(1, "Value") != 1.5) dynAppend(failed, "Data");

  // load() replaces the sheets in memory.
  in.addSheet("Scratch");
  if (!in.load() || in.hasSheet("Scratch")) dynAppend(failed, "load replaces sheets");

  remove(filename);
  return report(name, failed);
}

// loadSheet(): one sheet, the first one for "", replacing a sheet of that name.
bool excelWorkbookLoadSheetTest()
{
  dyn_string failed;
  string filename = tmpnam();

  shared_ptr<ExcelWorkbook> out = new ExcelWorkbook(filename);
  out.addSheet("Alpha").appendRow(makeMapping("k", "a"));
  out.addSheet("Beta").appendRow(makeMapping("k", "b"));

  if (!out.save())
  {
    dynAppend(failed, "save: " + out.getMessageText());
    return report("excelWorkbookLoadSheetTest", failed);
  }

  shared_ptr<ExcelWorkbook> in = new ExcelWorkbook(filename);
  shared_ptr<ExcelSheet> beta = in.loadSheet("Beta");

  if (beta == nullptr || beta.getValue(1, "k") != "b" || in.getSheetNames() != makeDynString("Beta"))
    dynAppend(failed, "loadSheet");

  shared_ptr<ExcelSheet> first = in.loadSheet();

  if (first == nullptr || first.getName() != "Alpha" || in.sheetCount() != 2) dynAppend(failed, "loadSheet first");

  beta.setValue(1, "k", "changed");
  in.loadSheet("Beta");

  if (in.sheetCount() != 2 || in.getSheet("Beta").getValue(1, "k") != "b") dynAppend(failed, "loadSheet replaces");
  if (beta.getValue(1, "k") != "changed") dynAppend(failed, "old sheet object untouched");

  if (in.loadSheet("Gamma") != nullptr || !in.hasMessages()) dynAppend(failed, "loadSheet missing sheet");
  if (strpos(in.getMessageText(), "Gamma") < 0) dynAppend(failed, "message names the sheet: " + in.getMessageText());

  remove(filename);
  return report("excelWorkbookLoadSheetTest", failed);
}

// setReadOptions(): without headers the keys are column numbers.
bool excelWorkbookReadOptionsTest()
{
  dyn_string failed;
  string filename = tmpnam();

  shared_ptr<ExcelWorkbook> out = new ExcelWorkbook(filename);
  out.addSheet("S").appendRow(makeMapping("Name", "Alice"));

  if (!out.save())
  {
    dynAppend(failed, "save: " + out.getMessageText());
    return report("excelWorkbookReadOptionsTest", failed);
  }

  shared_ptr<ExcelWorkbook> in = new ExcelWorkbook(filename);
  in.setReadOptions(TRUE, FALSE);
  shared_ptr<ExcelSheet> sheet = in.loadSheet("S");

  if (sheet == nullptr || sheet.rowCount() != 2 || sheet.getValue(1, 1) != "Name" || sheet.getValue(2, 1) != "Alice")
    dynAppend(failed, "keys are column numbers");

  remove(filename);
  return report("excelWorkbookReadOptionsTest", failed);
}

// Failures: missing file, unwritable path, invalid sheet name.
bool excelWorkbookErrorsTest()
{
  dyn_string failed;
  string missing = tmpnam();
  remove(missing);

  shared_ptr<ExcelWorkbook> book = new ExcelWorkbook(missing);

  if (book.exists()) dynAppend(failed, "exists");
  if (book.load() || !book.hasMessages()) dynAppend(failed, "load of a missing file");
  if (dynlen(book.readSheetNames()) != 0 || !book.hasMessages()) dynAppend(failed, "readSheetNames of a missing file");
  if (book.loadSheet() != nullptr) dynAppend(failed, "loadSheet of a missing file");

  // A directory that does not exist.
  shared_ptr<ExcelWorkbook> nowhere = new ExcelWorkbook(missing + "/sub/out.xlsx");
  nowhere.addSheet("S").appendRow(makeMapping("a", 1));
  if (nowhere.save() || !nowhere.hasMessages()) dynAppend(failed, "save to a missing directory");

  // An invalid sheet name fails the save and leaves the existing file alone.
  string filename = tmpnam();
  shared_ptr<ExcelWorkbook> good = new ExcelWorkbook(filename);
  good.addSheet("Good").appendRow(makeMapping("a", 1));
  if (!good.save()) dynAppend(failed, "save: " + good.getMessageText());

  shared_ptr<ExcelWorkbook> bad = new ExcelWorkbook(filename);
  bad.addSheet("a/b").appendRow(makeMapping("a", 2));
  if (bad.save() || !bad.hasMessages()) dynAppend(failed, "save with an invalid sheet name");
  if (good.readSheetNames() != makeDynString("Good")) dynAppend(failed, "file unchanged after a failed save");

  // A successful call clears the messages of the previous one.
  if (!good.load() || good.hasMessages()) dynAppend(failed, "messages cleared");

  remove(filename);
  return report("excelWorkbookErrorsTest", failed);
}

// The fixture: warnings on success, file order of the sheets.
bool excelWorkbookFixtureTest()
{
  string filename = getPath(DATA_REL_PATH, "CtrlExcelReaderFixture.xlsx");

  if (filename == "")
  {
    skipTest("excelWorkbookFixtureTest", "copy testdata/CtrlExcelReaderFixture.xlsx into the project's data directory");
    return TRUE;
  }

  dyn_string failed;
  shared_ptr<ExcelWorkbook> book = new ExcelWorkbook(filename);

  // Unrepresentable date and uncached formula: warnings, the load succeeds.
  if (!book.load()) dynAppend(failed, "load: " + book.getMessageText());
  if (dynlen(book.getMessages()) != 2) dynAppend(failed, "warnings " + dynlen(book.getMessages()));
  if (book.getSheetNames() != makeDynString("Fixture", "Implicit")) dynAppend(failed, "file order");

  shared_ptr<ExcelSheet> fixture = book.getSheet("Fixture");
  if (fixture == nullptr || dynlen(fixture.findRows("Label", "after")) != 1) dynAppend(failed, "Fixture rows");

  shared_ptr<ExcelSheet> implicit = book.getSheet("Implicit");
  if (implicit == nullptr || implicit.getColumn("Value") != makeDynAnytype(1, 2, 3, 4)) dynAppend(failed, "Implicit");

  return report("excelWorkbookFixtureTest", failed);
}
