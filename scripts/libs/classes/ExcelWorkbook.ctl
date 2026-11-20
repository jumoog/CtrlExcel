// $License: NOLICENSE
//--------------------------------------------------------------------------------
/**
  @file $relPath
  @copyright $copyright
  @author Kilian von Pflugk

  CTRL++ classes on top of the CtrlExcelReader extension:

  - ExcelWorkbook: one .xlsx file. load() / loadSheet() read sheets from the
    file into the workbook, save() writes all its sheets back; readSheetNames()
    lists the sheets in the file. Every file call keeps the messages of
    getLastError() (getMessages()), and uses the ...Async functions unless the
    call mode is ExcelCallMode::Blocking, so the other scripts of the manager
    keep running while a large file is read or written.
  - ExcelSheet: a named sheet in memory, a dyn_mapping of rows (header text, or
    the 1-based column number, -> value), with row, cell and column access.
    Row indexes are 1-based, like dyn indexes in CTRL.

  Usage:

    #uses "classes/ExcelWorkbook"

    shared_ptr<ExcelWorkbook> book = new ExcelWorkbook("C:/data/report.xlsx");
    if (!book.load())
      DebugTN("load failed", book.getMessageText());

    shared_ptr<ExcelSheet> people = book.getSheet("People");
    for (int i = 1; i <= people.rowCount(); i++)
      DebugTN(people.getValue(i, "Name"), people.getValue(i, "Age"));

    shared_ptr<ExcelWorkbook> out = new ExcelWorkbook("C:/data/out.xlsx");
    shared_ptr<ExcelSheet> sheet = out.addSheet("Result");
    sheet.appendRow(makeMapping("Name", "Alice", "Age", 30));
    if (!out.save())
      DebugTN("save failed", out.getMessageText());

  CTRL mappings do not keep the order in which keys were added, so the
  columns of a written sheet, and the sheets save() writes, come in mapping
  key order, as with excelWriteSheet / excelWriteFile. load() keeps the
  order of the file.
*/

#uses "CtrlExcelReader"

//--------------------------------------------------------------------------------
/// How ExcelWorkbook calls the extension.
enum ExcelCallMode
{
  /// ...Async functions: only the calling script waits (default).
  Async,
  /// Blocking functions: every script of the manager waits. A blocking call
  /// on a file another script's ...Async call is using fails with "busy".
  Blocking
};

//--------------------------------------------------------------------------------
/// A named sheet in memory. Rows are mappings from the column key (header
/// text, or the 1-based column number when the sheet was read without
/// headers) to the cell value; see excelReadSheet for the value types.
class ExcelSheet
{
  public ExcelSheet(string name = "", dyn_mapping rows = makeDynMapping())
  {
    m_name = name;
    m_rows = rows;
  }

  public string getName()
  {
    return m_name;
  }

  public void setName(string name)
  {
    m_name = name;
  }

  public int rowCount()
  {
    return dynlen(m_rows);
  }

  public dyn_mapping getRows()
  {
    return m_rows;
  }

  public void setRows(const dyn_mapping &rows)
  {
    m_rows = rows;
  }

  /// Row index (1-based), or an empty mapping if there is no such row.
  public mapping getRow(int index)
  {
    if (!isRowIndex(index))
      return makeMapping();

    return m_rows[index];
  }

  /// Replaces row index; index rowCount() + 1 appends. FALSE for other indexes.
  public bool setRow(int index, const mapping &row)
  {
    if (index == rowCount() + 1)
    {
      dynAppend(m_rows, row);
      return TRUE;
    }

    if (!isRowIndex(index))
      return FALSE;

    m_rows[index] = row;
    return TRUE;
  }

  public void appendRow(const mapping &row)
  {
    dynAppend(m_rows, row);
  }

  public bool removeRow(int index)
  {
    if (!isRowIndex(index))
      return FALSE;

    dynRemove(m_rows, index);
    return TRUE;
  }

  public void clear()
  {
    dynClear(m_rows);
  }

  /// Cell value, or "" if the row or the column does not exist (as for an
  /// empty cell).
  public anytype getValue(int row, anytype column)
  {
    if (!isRowIndex(row) || !mappingHasKey(m_rows[row], column))
      return "";

    return m_rows[row][column];
  }

  /// Sets a cell; row rowCount() + 1 appends a row. FALSE for other rows.
  public bool setValue(int row, anytype column, anytype value)
  {
    if (row == rowCount() + 1)
    {
      mapping newRow;
      newRow[column] = value;
      dynAppend(m_rows, newRow);
      return TRUE;
    }

    if (!isRowIndex(row))
      return FALSE;

    m_rows[row][column] = value;
    return TRUE;
  }

  /// Keys of all rows in order of first appearance (within a row, in mapping
  /// key order), i.e. the columns a write produces.
  public dyn_anytype getColumnNames()
  {
    dyn_anytype names;

    for (int r = 1; r <= dynlen(m_rows); r++)
    {
      for (int k = 1; k <= mappinglen(m_rows[r]); k++)
      {
        anytype key = mappingGetKey(m_rows[r], k);

        if (dynContains(names, key) <= 0)
          dynAppend(names, key);
      }
    }

    return names;
  }

  /// The column's value in every row ("" where a row lacks the column).
  public dyn_anytype getColumn(anytype column)
  {
    dyn_anytype values;

    for (int r = 1; r <= dynlen(m_rows); r++)
    {
      dynAppend(values, mappingHasKey(m_rows[r], column) ? m_rows[r][column] : "");
    }

    return values;
  }

  /// Rows whose column equals value.
  public dyn_mapping findRows(anytype column, anytype value)
  {
    dyn_mapping found;

    for (int r = 1; r <= dynlen(m_rows); r++)
    {
      if (mappingHasKey(m_rows[r], column) && m_rows[r][column] == value)
        dynAppend(found, m_rows[r]);
    }

    return found;
  }

  private bool isRowIndex(int index)
  {
    return index >= 1 && index <= dynlen(m_rows);
  }

  private string m_name;
  private dyn_mapping m_rows;
};

//--------------------------------------------------------------------------------
/// One .xlsx file and the sheets read from it or to be written to it.
class ExcelWorkbook
{
  public ExcelWorkbook(string filename)
  {
    m_filename = filename;
  }

  public string getFilename()
  {
    return m_filename;
  }

  public void setFilename(string filename)
  {
    m_filename = filename;
  }

  public ExcelCallMode getCallMode()
  {
    return m_mode;
  }

  public void setCallMode(ExcelCallMode mode)
  {
    m_mode = mode;
  }

  /// Options of later load() / loadSheet() calls; see excelReadSheet.
  public void setReadOptions(bool skipHiddenRows = TRUE, bool firstRowIsColumnNames = TRUE,
                             bool skipEmptyRows = TRUE)
  {
    m_skipHiddenRows = skipHiddenRows;
    m_firstRowIsColumnNames = firstRowIsColumnNames;
    m_skipEmptyRows = skipEmptyRows;
  }

  public bool exists()
  {
    return isfile(m_filename);
  }

  //------------------------------------------------------------------------------
  // File access. Each call replaces getMessages().

  /// Names of the worksheets in the file (empty on failure).
  public dyn_string readSheetNames()
  {
    dyn_string names;

    if (m_mode == ExcelCallMode::Async)
      excelGetSheetNamesAsync(m_filename, names);
    else
      names = excelGetSheetNames(m_filename);

    m_messages = getLastError();
    return names;
  }

  /// Replaces the workbook's sheets with all sheets of the file, in the
  /// file's order. FALSE if the file or one of its sheets could not be read;
  /// the sheets that could be read are kept. Messages on success are warnings
  /// (e.g. dates outside the range of CTRL time).
  public bool load()
  {
    dyn_string order = readSheetNames();
    dyn_errClass orderMessages = m_messages;

    if (dynlen(order) == 0 && dynlen(orderMessages) > 0)
      return FALSE;

    mapping sheets;

    if (m_mode == ExcelCallMode::Async)
      excelReadFileAsync(m_filename, sheets, m_skipHiddenRows, m_firstRowIsColumnNames, m_skipEmptyRows);
    else
      sheets = excelReadFile(m_filename, m_skipHiddenRows, m_firstRowIsColumnNames, m_skipEmptyRows);

    m_messages = getLastError();
    m_sheets.clear();

    // The file's order; names the project encoding cannot hold come back
    // suffixed (see excelReadFile) and go last.
    for (int i = 1; i <= dynlen(order); i++)
    {
      if (mappingHasKey(sheets, order[i]))
        m_sheets.append(new ExcelSheet(order[i], sheets[order[i]]));
    }

    for (int k = 1; k <= mappinglen(sheets); k++)
    {
      string name = mappingGetKey(sheets, k);

      if (dynContains(order, name) <= 0)
        m_sheets.append(new ExcelSheet(name, mappingGetValue(sheets, k)));
    }

    // excelReadFile leaves out the sheets it could not read.
    return m_sheets.count() >= dynlen(order);
  }

  /// Reads one sheet of the file ("" = the first) into the workbook,
  /// replacing a sheet of the same name. Returns it, or nullptr if it could
  /// not be read.
  public shared_ptr<ExcelSheet> loadSheet(string name = "")
  {
    if (name == "")
    {
      dyn_string names = readSheetNames();

      if (dynlen(names) == 0)
        return nullptr;

      name = names[1];
    }

    dyn_mapping rows;

    if (m_mode == ExcelCallMode::Async)
      excelReadSheetAsync(m_filename, name, rows, m_skipHiddenRows, m_firstRowIsColumnNames, m_skipEmptyRows);
    else
      rows = excelReadSheet(m_filename, name, m_skipHiddenRows, m_firstRowIsColumnNames, m_skipEmptyRows);

    m_messages = getLastError();

    // Read functions return nothing on failure; an empty sheet has no messages.
    if (dynlen(rows) == 0 && dynlen(m_messages) > 0)
      return nullptr;

    shared_ptr<ExcelSheet> sheet = new ExcelSheet(name, rows);
    int index = sheetIndex(name);

    // Not m_sheets[index] = sheet: assigning to a vector element copies into
    // the object it points to, which would change sheets the caller holds.
    if (index >= 0)
    {
      m_sheets.removeAt(index);
      m_sheets.insertAt(index, sheet);
    }
    else
    {
      m_sheets.append(sheet);
    }

    return sheet;
  }

  /// Writes all sheets to the file, replacing it (a workbook without sheets
  /// writes one empty sheet "Sheet1"). A failed write leaves an existing file
  /// unchanged. The file's sheet order is the mapping order of excelWriteFile,
  /// not necessarily the workbook's.
  public bool save()
  {
    mapping data;

    for (int i = 0; i < m_sheets.count(); i++)
    {
      data[m_sheets.at(i).getName()] = m_sheets.at(i).getRows();
    }

    bool ok = FALSE;

    if (m_mode == ExcelCallMode::Async)
      excelWriteFileAsync(m_filename, data, ok);
    else
      ok = excelWriteFile(m_filename, data);

    m_messages = getLastError();
    return ok;
  }

  //------------------------------------------------------------------------------
  // Sheets in memory. Names are compared exactly (Excel ignores case, which
  // save() checks).

  public int sheetCount()
  {
    return m_sheets.count();
  }

  public dyn_string getSheetNames()
  {
    dyn_string names;

    for (int i = 0; i < m_sheets.count(); i++)
    {
      dynAppend(names, m_sheets.at(i).getName());
    }

    return names;
  }

  public bool hasSheet(string name)
  {
    return sheetIndex(name) >= 0;
  }

  /// The sheet, or nullptr if the workbook has none of that name.
  public shared_ptr<ExcelSheet> getSheet(string name)
  {
    int index = sheetIndex(name);

    if (index < 0)
      return nullptr;

    return m_sheets.at(index);
  }

  /// The sheet of that name, appended empty if the workbook has none yet.
  public shared_ptr<ExcelSheet> addSheet(string name)
  {
    int index = sheetIndex(name);

    if (index >= 0)
      return m_sheets.at(index);

    shared_ptr<ExcelSheet> sheet = new ExcelSheet(name);
    m_sheets.append(sheet);
    return sheet;
  }

  public bool removeSheet(string name)
  {
    int index = sheetIndex(name);

    if (index < 0)
      return FALSE;

    m_sheets.removeAt(index);
    return TRUE;
  }

  //------------------------------------------------------------------------------
  // Messages of the last file call: errors, or warnings when it succeeded.

  public dyn_errClass getMessages()
  {
    return m_messages;
  }

  public bool hasMessages()
  {
    return dynlen(m_messages) > 0;
  }

  /// The messages' texts, one per line.
  public string getMessageText()
  {
    string text;

    for (int i = 1; i <= dynlen(m_messages); i++)
    {
      if (i > 1)
        text += "\n";
      text += getErrorText(m_messages[i]);
    }

    return text;
  }

  private int sheetIndex(string name)
  {
    for (int i = 0; i < m_sheets.count(); i++)
    {
      if (m_sheets.at(i).getName() == name)
        return i;
    }

    return -1;
  }

  private string m_filename;
  private ExcelCallMode m_mode = ExcelCallMode::Async;
  private bool m_skipHiddenRows = TRUE;
  private bool m_firstRowIsColumnNames = TRUE;
  private bool m_skipEmptyRows = TRUE;
  private vector< shared_ptr<ExcelSheet> > m_sheets;
  private dyn_errClass m_messages;
};
