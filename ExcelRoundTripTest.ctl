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

//--------------------------------------------------------------------------------
/**
*/
void main()
{
  excelRoundTripTestSingle();
  excelRoundTripTestFile();
  excelRoundTripTestIntKeys();
  excelFailedWriteKeepsFileTest();
  excelReadErrorReportedTest();
  excelLongValueTest();
  excelEmptyHeaderTest();
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
    DebugTN(context + ": unexpected row count", dynlen(rows));
    return false;
  }

  dyn_string missingKeys;
  dyn_string keysToCheck = makeDynString("Name", "Age", "Active", "Time");

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

  bool pass = rows[1]["Name"]   == "Alice"
              && rows[1]["Age"]    == 30
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
