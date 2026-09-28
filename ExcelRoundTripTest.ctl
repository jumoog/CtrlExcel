// $License: NOLICENSE
//--------------------------------------------------------------------------------
/**
  @file $relPath
  @copyright $copyright
  @author Kilian von Pflugk
*/

//--------------------------------------------------------------------------------
// Libraries used (#uses)
#uses "CtrlExcelReader"


//--------------------------------------------------------------------------------
// Variables and Constants

// Results collected by recordTest() / skipTest() for the summary in main().
int g_passedTests;
dyn_string g_failedTests;
dyn_string g_skippedTests;

// excelAsyncDoesNotBlockTest: ticks counted by a parallel CTRL thread.
int g_ticks;
bool g_ticking;

// excelAsyncConcurrencyTest: parallel CTRL threads report here.
int g_concurrentDone;
dyn_string g_concurrentFailures;

//--------------------------------------------------------------------------------
/**
*/
void main()
{
  recordTest("excelRoundTripTestSingle", excelRoundTripTestSingle());
  recordTest("excelRoundTripTestFile", excelRoundTripTestFile());
  recordTest("excelRoundTripTestIntKeys", excelRoundTripTestIntKeys());
  recordTest("excelFailedWriteKeepsFileTest", excelFailedWriteKeepsFileTest());
  recordTest("excelReadErrorReportedTest", excelReadErrorReportedTest());
  recordTest("excelLongValueTest", excelLongValueTest());
  recordTest("excelEmptyHeaderTest", excelEmptyHeaderTest());
  recordTest("excelSheetNameValidationTest", excelSheetNameValidationTest());
  recordTest("excelControlCharsTest", excelControlCharsTest());
  recordTest("excelTextTooLongTest", excelTextTooLongTest());
  recordTest("excelUnionKeysTest", excelUnionKeysTest());
  recordTest("excelMidnightDateTest", excelMidnightDateTest());
  recordTest("excelLargeFloatTest", excelLargeFloatTest());
  recordTest("excelWideTextTest", excelWideTextTest());
  recordTest("excelTooManyColumnsTest", excelTooManyColumnsTest());
  recordTest("excelEmptySheetNameTest", excelEmptySheetNameTest());
  recordTest("excelEmptyWriteFileTest", excelEmptyWriteFileTest());
  recordTest("excelFixtureTest", excelFixtureTest());
  recordTest("excelAsyncRoundTripTest", excelAsyncRoundTripTest());
  recordTest("excelAsyncDoesNotBlockTest", excelAsyncDoesNotBlockTest());
  recordTest("excelAsyncConcurrencyTest", excelAsyncConcurrencyTest());
  // Last: passes a dyn where the signature declares a mapping.
  recordTest("excelWriteFileWrongTypeTest", excelWriteFileWrongTypeTest());

  DebugTN("ExcelRoundTripTest summary",
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

// Returns a temporary filename, or "" on failure.
string getTempFile(string context)
{
  string filename = tmpnam();

  if (filename == "")
  {
    DebugTN(context + ": tmpnam failed");
  }

  return filename;
}

// Builds the shared two-row test dataset.
// t1/t2 are set to the time values written, so callers can compare on read-back.
dyn_anytype buildTestRows(time &t1, time &t2)
{
  // Milliseconds are part of the Excel serial and must survive the round-trip.
  t1 = getCurrentTime();
  t2 = makeTime(2026, 1, 1, 1, 1, 1, 999);

  mapping row1;
  row1["Name"]   = "Alice";
  row1["Age"]    = 30;
  row1["Score"]  = 95.5;
  row1["Active"] = TRUE;
  row1["Time"]   = t1;

  mapping row2;
  row2["Name"]   = "Bob";
  row2["Age"]    = 25;
  row2["Score"]  = 87.0;
  row2["Active"] = FALSE;
  row2["Time"]   = t2;

  dyn_anytype rows;
  dynAppend(rows, row1);
  dynAppend(rows, row2);
  return rows;
}

// Verifies read-back rows against the expected values written by buildTestRows().
bool checkRows(dyn_mapping rows, time t1, time t2, string context = "checkRows")
{
  if (dynlen(rows) != 2)
  {
    // getLastError() still holds the errors of the read that produced rows.
    DebugTN(context + ": unexpected row count", dynlen(rows), getLastError());
    return false;
  }

  dyn_string missingKeys;
  dyn_string keysToCheck = makeDynString("Name", "Age", "Score", "Active", "Time");

  for (int row = 1; row <= 2; row++)
  {
    for (int k = 1; k <= dynlen(keysToCheck); k++)
    {
      if (!mappingHasKey(rows[row], keysToCheck[k]))
      {
        dynAppend(missingKeys, "row" + row + "." + keysToCheck[k]);
      }
    }
  }

  if (dynlen(missingKeys) > 0)
  {
    DebugTN(context + ": missing keys", missingKeys);
    return false;
  }

  // Excel has no int/float distinction: 87.0 is stored as "87" and comes
  // back as int, 95.5 as float.
  bool pass = rows[1]["Name"]   == "Alice"
              && rows[1]["Age"]    == 30
              && getType(rows[1]["Score"]) == FLOAT_VAR && rows[1]["Score"] == 95.5
              && getType(rows[2]["Score"]) == INT_VAR   && rows[2]["Score"] == 87
              && rows[1]["Active"] == TRUE
              && rows[1]["Time"]   == t1
              && rows[2]["Name"]   == "Bob"
              && rows[2]["Age"]    == 25
              && rows[2]["Active"] == FALSE
              && rows[2]["Time"]   == t2;

  if (!pass)
  {
    DebugTN(context + ": value mismatch — read-back data", rows, "expected t1", t1, "expected t2", t2);
  }

  return pass;
}

// Round-trip test for excelWriteSheet / excelGetSheetNames / excelReadSheet.
bool excelRoundTripTestSingle(string filename = "")
{
  if (filename == "")
  {
    filename = getTempFile("excelRoundTripTestSingle");

    if (filename == "") return FALSE;
  }

  time t1, t2;
  dyn_anytype rows = buildTestRows(t1, t2);

  if (!excelWriteSheet(filename, "People", rows))
  {
    DebugTN("excelRoundTripTestSingle: excelWriteSheet failed", filename);
    return FALSE;
  }

  if (!excelGetSheetNames(filename).contains("People"))
  {
    DebugTN("excelRoundTripTestSingle: excelGetSheetNames failed", filename);
    return FALSE;
  }

  bool pass = checkRows(excelReadSheet(filename, "People"), t1, t2, "excelRoundTripTestSingle");
  DebugTN("excelRoundTripTestSingle", "file", filename, "pass", pass);
  remove(filename);
  return pass;
}

// Round-trip test for excelWriteFile / excelReadFile (multi-sheet API).
bool excelRoundTripTestFile(string filename = "")
{
  if (filename == "")
  {
    filename = getTempFile("excelRoundTripTestFile");

    if (filename == "") return FALSE;
  }

  time t1, t2;
  mapping data;
  data["People"] = buildTestRows(t1, t2);

  if (!excelWriteFile(filename, data))
  {
    DebugTN("excelRoundTripTestFile: excelWriteFile failed", filename);
    return FALSE;
  }

  mapping back = excelReadFile(filename);

  if (!mappingHasKey(back, "People"))
  {
    DebugTN("excelRoundTripTestFile: sheet 'People' missing from read-back", mappingKeys(back));
    return false;
  }

  bool pass = checkRows(back["People"], t1, t2, "excelRoundTripTestFile");
  DebugTN("excelRoundTripTestFile", "file", filename, "pass", pass);
  remove(filename);
  return pass;
}

// Round-trip for rows keyed by integers (the shape excelReadSheet returns with
// firstRowIsColumnNames = false): values must be written, not blank cells.
bool excelRoundTripTestIntKeys(string filename = "")
{
  if (filename == "")
  {
    filename = getTempFile("excelRoundTripTestIntKeys");

    if (filename == "") return FALSE;
  }

  dyn_anytype rows;
  dynAppend(rows, makeMapping(1, "Alice", 2, 30));
  dynAppend(rows, makeMapping(1, "Bob",   2, 25));

  if (!excelWriteSheet(filename, "Data", rows))
  {
    DebugTN("excelRoundTripTestIntKeys: excelWriteSheet failed", filename, getLastError());
    return FALSE;
  }

  // Headers are the stringified keys "1" and "2".
  dyn_mapping back = excelReadSheet(filename, "Data");

  bool pass = dynlen(back) == 2
              && back[1]["1"] == "Alice" && back[1]["2"] == 30
              && back[2]["1"] == "Bob"   && back[2]["2"] == 25;

  if (!pass)
  {
    DebugTN("excelRoundTripTestIntKeys: value mismatch — read-back data", back);
  }

  DebugTN("excelRoundTripTestIntKeys", "file", filename, "pass", pass);
  remove(filename);
  return pass;
}

// A write with an invalid row must return FALSE, report an error and leave the
// previously written file untouched.
bool excelFailedWriteKeepsFileTest(string filename = "")
{
  if (filename == "")
  {
    filename = getTempFile("excelFailedWriteKeepsFileTest");

    if (filename == "") return FALSE;
  }

  time t1, t2;
  dyn_anytype rows = buildTestRows(t1, t2);

  if (!excelWriteSheet(filename, "People", rows))
  {
    DebugTN("excelFailedWriteKeepsFileTest: initial excelWriteSheet failed", filename, getLastError());
    return FALSE;
  }

  dyn_anytype badRows = rows;
  dynAppend(badRows, "not a mapping");

  bool writeOk = excelWriteSheet(filename, "Other", badRows);
  int errCount = dynlen(getLastError());

  bool pass = !writeOk
              && errCount > 0
              && checkRows(excelReadSheet(filename, "People"), t1, t2, "excelFailedWriteKeepsFileTest");

  DebugTN("excelFailedWriteKeepsFileTest", "file", filename, "writeOk", writeOk, "errors", errCount, "pass", pass);
  remove(filename);
  return pass;
}

// Reading a missing file returns an empty result and reports an error.
bool excelReadErrorReportedTest()
{
  string filename = getTempFile("excelReadErrorReportedTest");

  if (filename == "") return FALSE;

  remove(filename); // make sure it does not exist

  dyn_mapping rows = excelReadSheet(filename, "People");
  int errCount = dynlen(getLastError());

  bool pass = dynlen(rows) == 0 && errCount > 0;
  DebugTN("excelReadErrorReportedTest", "file", filename, "errors", errCount, "pass", pass);
  return pass;
}

// Integers beyond 32 bits must come back as long, not wrapped to int.
bool excelLongValueTest(string filename = "")
{
  if (filename == "")
  {
    filename = getTempFile("excelLongValueTest");

    if (filename == "") return FALSE;
  }

  long big = 50000;
  big = big * 100000; // 5000000000, beyond the int range

  dyn_anytype rows;
  dynAppend(rows, makeMapping("Big", big));

  if (!excelWriteSheet(filename, "Data", rows))
  {
    DebugTN("excelLongValueTest: excelWriteSheet failed", filename, getLastError());
    return FALSE;
  }

  dyn_mapping back = excelReadSheet(filename, "Data");

  bool pass = dynlen(back) == 1
              && getType(back[1]["Big"]) == LONG_VAR
              && back[1]["Big"] == big;

  if (!pass)
  {
    DebugTN("excelLongValueTest: value mismatch — read-back data", back);
  }

  DebugTN("excelLongValueTest", "file", filename, "pass", pass);
  remove(filename);
  return pass;
}

// An empty header cell falls back to the 1-based column number as key.
bool excelEmptyHeaderTest(string filename = "")
{
  if (filename == "")
  {
    filename = getTempFile("excelEmptyHeaderTest");

    if (filename == "") return FALSE;
  }

  dyn_anytype rows;
  dynAppend(rows, makeMapping("", "first", "Name", "Alice"));

  if (!excelWriteSheet(filename, "Data", rows))
  {
    DebugTN("excelEmptyHeaderTest: excelWriteSheet failed", filename, getLastError());
    return FALSE;
  }

  dyn_mapping back = excelReadSheet(filename, "Data");

  // Mapping key order decides the column, so look the empty-header column up
  // by whichever integer key was produced.
  bool pass = dynlen(back) == 1 && back[1]["Name"] == "Alice";
  int intKeys = 0;

  if (pass)
  {
    dyn_anytype keys = mappingKeys(back[1]);

    for (int i = 1; i <= dynlen(keys); i++)
    {
      if (getType(keys[i]) == INT_VAR && back[1][keys[i]] == "first")
        intKeys++;
    }
  }

  pass = pass && intKeys == 1;

  if (!pass)
  {
    DebugTN("excelEmptyHeaderTest: unexpected read-back data", back);
  }

  DebugTN("excelEmptyHeaderTest", "file", filename, "pass", pass);
  remove(filename);
  return pass;
}

// Writes rows to a temp file with excelWriteSheet and reads them back.
// Returns FALSE (and logs) if the write fails.
bool writeAndReadBack(string context, dyn_anytype rows, dyn_mapping &back)
{
  string filename = getTempFile(context);

  if (filename == "") return FALSE;

  bool ok = excelWriteSheet(filename, "Data", rows);

  if (!ok)
  {
    DebugTN(context + ": excelWriteSheet failed", filename, getLastError());
  }
  else
  {
    back = excelReadSheet(filename, "Data");
  }

  remove(filename);
  return ok;
}

// Sheet names Excel rejects must fail the write with an error instead of
// producing a corrupt workbook.
bool excelSheetNameValidationTest()
{
  string filename = getTempFile("excelSheetNameValidationTest");

  if (filename == "") return FALSE;

  dyn_anytype rows = makeDynAnytype(makeMapping("A", 1));
  mapping sameIgnoringCase = makeMapping("Data", rows, "data", rows);
  mapping sameIgnoringUmlautCase = makeMapping("Übersicht", rows, "übersicht", rows);

  bool slashRejected  = !excelWriteSheet(filename, "2026/09", rows) && dynlen(getLastError()) > 0;
  bool longRejected   = !excelWriteSheet(filename, "abcdefghijklmnopqrstuvwxyz123456", rows) && dynlen(getLastError()) > 0;
  bool caseRejected   = !excelWriteFile(filename, sameIgnoringCase) && dynlen(getLastError()) > 0;
  bool umlautRejected = !excelWriteFile(filename, sameIgnoringUmlautCase) && dynlen(getLastError()) > 0;
  bool validAccepted  = excelWriteSheet(filename, "Data 2026-09", rows);

  bool pass = slashRejected && longRejected && caseRejected && umlautRejected && validAccepted;
  DebugTN("excelSheetNameValidationTest", "slash", slashRejected, "long", longRejected,
          "case", caseRejected, "umlautCase", umlautRejected, "valid", validAccepted, "pass", pass);
  remove(filename);
  return pass;
}

// Control characters and CR, which XML cannot hold or normalises, are stored
// as Excel's _xHHHH_ escapes and round-trip exactly, as does a literal escape.
bool excelControlCharsTest()
{
  string text;
  sprintf(text, "a%cb\tc\r\nd _x0041_", 1);
  dyn_mapping back;

  bool pass = writeAndReadBack("excelControlCharsTest", makeDynAnytype(makeMapping("Text", text)), back)
              && dynlen(back) == 1
              && back[1]["Text"] == text;

  if (!pass)
  {
    DebugTN("excelControlCharsTest: mismatch — read-back data", back);
  }

  DebugTN("excelControlCharsTest", "pass", pass);
  return pass;
}

// Text beyond Excel's 32767-character cell limit fails the write.
bool excelTextTooLongTest()
{
  string filename = getTempFile("excelTextTooLongTest");

  if (filename == "") return FALSE;

  string text = "x";

  while (strlen(text) <= 32767)
  {
    text += text;
  }

  bool writeOk = excelWriteSheet(filename, "Data", makeDynAnytype(makeMapping("Text", text)));
  int errCount = dynlen(getLastError());

  bool pass = !writeOk && errCount > 0;
  DebugTN("excelTextTooLongTest", "length", strlen(text), "writeOk", writeOk, "errors", errCount, "pass", pass);
  remove(filename);
  return pass;
}

// Keys that only appear in later rows still become columns.
bool excelUnionKeysTest()
{
  dyn_anytype rows;
  dynAppend(rows, makeMapping("Name", "Alice"));
  dynAppend(rows, makeMapping("Name", "Bob", "Comment", "late key"));
  dyn_mapping back;

  bool pass = writeAndReadBack("excelUnionKeysTest", rows, back)
              && dynlen(back) == 2
              && mappingHasKey(back[1], "Comment") && back[1]["Comment"] == ""
              && back[2]["Comment"] == "late key";

  if (!pass)
  {
    DebugTN("excelUnionKeysTest: mismatch — read-back data", back);
  }

  DebugTN("excelUnionKeysTest", "pass", pass);
  return pass;
}

// excelWriteFile with rows instead of a sheet mapping must fail, not return
// TRUE without writing.
bool excelWriteFileWrongTypeTest()
{
  string filename = getTempFile("excelWriteFileWrongTypeTest");

  if (filename == "") return FALSE;

  remove(filename);

  dyn_anytype rows = makeDynAnytype(makeMapping("A", 1));
  bool writeOk = excelWriteFile(filename, rows);
  int errCount = dynlen(getLastError());

  bool pass = !writeOk && errCount > 0 && !isfile(filename);
  DebugTN("excelWriteFileWrongTypeTest", "writeOk", writeOk, "errors", errCount, "pass", pass);
  remove(filename);
  return pass;
}

// A time at exactly midnight has a whole-number serial, which OpenXLSX types
// as integer; it must still come back as a time, not an int.
bool excelMidnightDateTest()
{
  time t = makeTime(2026, 1, 1);
  dyn_mapping back;

  bool pass = writeAndReadBack("excelMidnightDateTest", makeDynAnytype(makeMapping("Date", t)), back)
              && dynlen(back) == 1
              && getType(back[1]["Date"]) == TIME_VAR
              && back[1]["Date"] == t;

  if (!pass)
  {
    DebugTN("excelMidnightDateTest: mismatch — read-back data", back, "expected", t);
  }

  DebugTN("excelMidnightDateTest", "pass", pass);
  return pass;
}

// Large floats are stored in exponent notation ("1e+20"); OpenXLSX's integer
// parsing would read that as 1.
bool excelLargeFloatTest()
{
  float big = 1e20;
  float negative = -3e18;
  dyn_mapping back;

  bool pass = writeAndReadBack("excelLargeFloatTest", makeDynAnytype(makeMapping("Big", big, "Negative", negative)), back)
              && dynlen(back) == 1
              && back[1]["Big"] == big
              && back[1]["Negative"] == negative;

  if (!pass)
  {
    DebugTN("excelLargeFloatTest: mismatch — read-back data", back);
  }

  DebugTN("excelLargeFloatTest", "pass", pass);
  return pass;
}

// Long texts round-trip; the column width is capped at Excel's maximum of 255
// (checked by opening the file in Excel, not here).
bool excelWideTextTest()
{
  string text = "0123456789";

  while (strlen(text) < 5000)
  {
    text += text;
  }

  dyn_mapping back;

  bool pass = writeAndReadBack("excelWideTextTest", makeDynAnytype(makeMapping("Text", text)), back)
              && dynlen(back) == 1
              && back[1]["Text"] == text;

  DebugTN("excelWideTextTest", "length", strlen(text), "pass", pass);
  return pass;
}

// More columns than Excel's 16384 must fail the write instead of producing
// a corrupt file.
bool excelTooManyColumnsTest()
{
  string filename = getTempFile("excelTooManyColumnsTest");

  if (filename == "") return FALSE;

  mapping row;

  for (int i = 1; i <= 16385; i++)
  {
    row[i] = i;
  }

  bool writeOk = excelWriteSheet(filename, "Data", makeDynAnytype(row));
  int errCount = dynlen(getLastError());

  bool pass = !writeOk && errCount > 0;
  DebugTN("excelTooManyColumnsTest", "writeOk", writeOk, "errors", errCount, "pass", pass);
  remove(filename);
  return pass;
}

// An empty sheet name reads the first worksheet.
bool excelEmptySheetNameTest()
{
  string filename = getTempFile("excelEmptySheetNameTest");

  if (filename == "") return FALSE;

  time t1, t2;

  if (!excelWriteSheet(filename, "People", buildTestRows(t1, t2)))
  {
    DebugTN("excelEmptySheetNameTest: excelWriteSheet failed", filename, getLastError());
    return FALSE;
  }

  // Also with every optional argument, incl. skipEmptyRows = FALSE.
  bool pass = checkRows(excelReadSheet(filename, ""), t1, t2, "excelEmptySheetNameTest")
              && checkRows(excelReadSheet(filename, "", FALSE, TRUE, FALSE), t1, t2, "excelEmptySheetNameTest (all args)");
  DebugTN("excelEmptySheetNameTest", "pass", pass);
  remove(filename);
  return pass;
}

// Reads testdata/CtrlExcelReaderFixture.xlsx (generated by
// testdata/make_fixture.py in Excel's own XML conventions) to cover cases the
// extension's writer cannot produce. Copy the file into the project's data
// directory; the test is skipped if it is missing.
bool excelFixtureTest()
{
  string filename = getPath(DATA_REL_PATH, "CtrlExcelReaderFixture.xlsx");

  if (filename == "")
  {
    skipTest("excelFixtureTest", "copy testdata/CtrlExcelReaderFixture.xlsx into the project's data directory");
    return TRUE;
  }

  dyn_mapping rows = excelReadSheet(filename, "Fixture");
  int warnings = dynlen(getLastError()); // unrepresentable date + uncached formula

  mapping byLabel;

  for (int i = 1; i <= dynlen(rows); i++)
  {
    byLabel[rows[i]["Label"]] = rows[i]["Value"];
  }

  dyn_string failed;

  // Formatted empty row 11 is skipped by default, kept with skipEmptyRows = FALSE.
  if (dynlen(rows) != 10) dynAppend(failed, "row count " + dynlen(rows));
  if (dynlen(excelReadSheet(filename, "Fixture", TRUE, TRUE, FALSE)) != 11) dynAppend(failed, "row count with empty rows");
  if (warnings != 2) dynAppend(failed, "warnings " + warnings);

  // 1960 is outside CTRL time: the Excel serial as float.
  if (getType(byLabel["pre1970"]) != FLOAT_VAR || byLabel["pre1970"] != 22037) dynAppend(failed, "pre1970");
  // Whole-number date serial (typed integer by OpenXLSX).
  if (getType(byLabel["dateInt"]) != TIME_VAR || byLabel["dateInt"] != makeTime(2026, 1, 1)) dynAppend(failed, "dateInt");
  // [h]:mm is a duration, not a date.
  if (getType(byLabel["duration"]) != FLOAT_VAR || byLabel["duration"] != 1.5) dynAppend(failed, "duration");
  if (byLabel["error"] != "#DIV/0!") dynAppend(failed, "error");
  if (byLabel["uncached"] != "") dynAppend(failed, "uncached");
  if (byLabel["crlf"] != "a\r\nb") dynAppend(failed, "crlf");
  // Printable _xHHHH_ stays literal.
  if (byLabel["literal"] != "AB_x0041_7") dynAppend(failed, "literal");
  if (byLabel["exponent"] != 1e20) dynAppend(failed, "exponent");
  if (getType(byLabel["bigint"]) != FLOAT_VAR || byLabel["bigint"] < 1.2e19) dynAppend(failed, "bigint");
  if (byLabel["after"] != "last") dynAppend(failed, "after");

  bool pass = dynlen(failed) == 0;

  if (!pass)
  {
    DebugTN("excelFixtureTest: failed checks", failed, "read-back data", rows);
  }

  DebugTN("excelFixtureTest", "pass", pass);
  return pass;
}

// excelWriteFile with an empty mapping writes a workbook with one empty sheet
// (like excelWriteSheet with empty data) instead of returning TRUE without a
// file.
bool excelEmptyWriteFileTest()
{
  string filename = getTempFile("excelEmptyWriteFileTest");

  if (filename == "") return FALSE;

  remove(filename);

  mapping noSheets;
  bool writeOk = excelWriteFile(filename, noSheets);
  bool exists = isfile(filename);
  dyn_string sheets = excelGetSheetNames(filename);

  bool pass = writeOk && exists && dynlen(sheets) == 1 && sheets[1] == "Sheet1"
              && dynlen(excelReadSheet(filename, "Sheet1")) == 0;
  DebugTN("excelEmptyWriteFileTest", "writeOk", writeOk, "exists", exists, "sheets", sheets, "pass", pass);
  remove(filename);
  return pass;
}

// The *Async variants deliver through reference parameters once the call
// returns; errors arrive in getLastError() as for the blocking functions.
bool excelAsyncRoundTripTest()
{
  string filename = getTempFile("excelAsyncRoundTripTest");
  string filename2 = getTempFile("excelAsyncRoundTripTest2");

  if (filename == "" || filename2 == "") return FALSE;

  dyn_string failed;
  time t1, t2;
  dyn_anytype rows = buildTestRows(t1, t2);

  bool ok;
  if (excelWriteSheetAsync(filename, "People", rows, ok) != 0 || !ok)
    dynAppend(failed, "excelWriteSheetAsync");

  dyn_string names;
  if (excelGetSheetNamesAsync(filename, names) != 0 || dynlen(names) != 1 || names[1] != "People")
    dynAppend(failed, "excelGetSheetNamesAsync");

  dyn_mapping back;
  if (excelReadSheetAsync(filename, "People", back) != 0
      || !checkRows(back, t1, t2, "excelAsyncRoundTripTest (excelReadSheetAsync)"))
    dynAppend(failed, "excelReadSheetAsync");

  mapping sheets;
  if (excelReadFileAsync(filename, sheets) != 0 || !mappingHasKey(sheets, "People")
      || !checkRows(sheets["People"], t1, t2, "excelAsyncRoundTripTest (excelReadFileAsync)"))
    dynAppend(failed, "excelReadFileAsync");

  bool ok2;
  if (excelWriteFileAsync(filename2, makeMapping("People", rows), ok2) != 0 || !ok2
      || !checkRows(excelReadSheet(filename2, "People"), t1, t2, "excelAsyncRoundTripTest (excelWriteFileAsync)"))
    dynAppend(failed, "excelWriteFileAsync");

  // Errors: a missing file yields no rows and an error after the call.
  remove(filename2);
  dyn_mapping missing;
  int started = excelReadSheetAsync(filename2, "People", missing);
  dyn_errClass asyncErrors = getLastError();
  if (started != 0 || dynlen(missing) != 0 || dynlen(asyncErrors) == 0)
  {
    dynAppend(failed, "error delivery");
    DebugTN("excelAsyncRoundTripTest: error delivery", "started", started,
            "rows", dynlen(missing), "errors", asyncErrors);
  }

  bool pass = dynlen(failed) == 0;
  DebugTN("excelAsyncRoundTripTest", "failed", failed, "pass", pass);
  remove(filename);
  remove(filename2);
  return pass;
}

void tickThread()
{
  while (g_ticking)
  {
    g_ticks++;
    delay(0, 10);
  }
}

// While an *Async read of a large sheet runs, other CTRL threads keep
// running (a blocking read would freeze the whole manager).
bool excelAsyncDoesNotBlockTest()
{
  string filename = getTempFile("excelAsyncDoesNotBlockTest");

  if (filename == "") return FALSE;

  dyn_anytype rows;

  for (int i = 1; i <= 20000; i++)
  {
    dynAppend(rows, makeMapping("Id", i, "Name", "row " + i, "Value", i * 1.5, "Flag", i % 2 == 0));
  }

  if (!excelWriteSheet(filename, "Big", rows))
  {
    DebugTN("excelAsyncDoesNotBlockTest: excelWriteSheet failed", getLastError());
    remove(filename);
    return FALSE;
  }

  g_ticks = 0;
  g_ticking = TRUE;
  int tid = startThread("tickThread");
  delay(0, 50);

  int ticksBefore = g_ticks;
  time start = getCurrentTime();
  dyn_mapping back;
  int started = excelReadSheetAsync(filename, "Big", back);
  float seconds = getCurrentTime() - start;
  int ticksDuring = g_ticks - ticksBefore;

  g_ticking = FALSE;
  delay(0, 50);

  // Expect roughly one tick per 10 ms of reading; require a clear majority
  // of that so a slow machine does not fail the test.
  int expected = seconds * 100;
  bool pass = started == 0 && dynlen(back) == 20000 && back[20000]["Id"] == 20000
              && ticksDuring >= expected / 2 && ticksDuring >= 2;
  DebugTN("excelAsyncDoesNotBlockTest", "read seconds", seconds, "ticks during read", ticksDuring,
          "rows", dynlen(back), "pass", pass);
  remove(filename);
  return pass;
}

void concurrentWriter(string filename, int writer)
{
  dyn_anytype rows;

  for (int i = 1; i <= 2000; i++)
  {
    dynAppend(rows, makeMapping("Writer", writer, "I", i));
  }

  bool ok;
  excelWriteSheetAsync(filename, "Data", rows, ok);

  if (!ok)
    dynAppend(g_concurrentFailures, "writer " + writer);

  g_concurrentDone++;
}

void concurrentReader(string filename, int reader)
{
  dyn_mapping rows;
  excelReadSheetAsync(filename, "Data", rows);

  // Each read must see one complete write, never a mix or a partial file.
  if (dynlen(rows) != 2000 || rows[1]["Writer"] != rows[2000]["Writer"])
    dynAppend(g_concurrentFailures, "reader " + reader + " saw " + dynlen(rows) + " rows");

  g_concurrentDone++;
}

// Several scripts writing and reading the same file through the *Async
// variants at once: operations on one path are serialised, so the file is
// never corrupted and every read sees a complete workbook.
bool excelAsyncConcurrencyTest()
{
  string filename = getTempFile("excelAsyncConcurrencyTest");

  if (filename == "") return FALSE;

  g_concurrentDone = 0;
  g_concurrentFailures = makeDynString();

  // Seed the file so early readers find a complete workbook.
  concurrentWriter(filename, 0);

  for (int i = 1; i <= 3; i++)
  {
    startThread("concurrentWriter", filename, i);
    startThread("concurrentReader", filename, i);
  }

  time deadline = getCurrentTime() + 60;

  while (g_concurrentDone < 7 && getCurrentTime() < deadline)
  {
    delay(0, 20);
  }

  dyn_mapping back = excelReadSheet(filename, "Data");
  bool finalOk = dynlen(back) == 2000 && back[1]["Writer"] == back[2000]["Writer"];

  bool pass = g_concurrentDone == 7 && dynlen(g_concurrentFailures) == 0 && finalOk;
  DebugTN("excelAsyncConcurrencyTest", "finished", g_concurrentDone, "failures", g_concurrentFailures,
          "final rows", dynlen(back), "pass", pass);
  remove(filename);
  return pass;
}
