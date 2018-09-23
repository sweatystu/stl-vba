Option Explicit

' Dependencies:
'   - cCustomErrors class

' Private Variables
Private vTbl As ListObject
Private vExistingRows As Long
Private vShowFilter As Boolean

' Public Procedures
Sub ClearTableContents()
    ' Description: Clears the contents of the whole table (not headers)
    '               Doesn't remove rows, just the contents
    ' Dependencies:
    '   - Me.lo()
    ' Inputs: None
    ' Outputs: None
    On Error GoTo ErrorHandle
    Me.lo.DataBodyRange.ClearContents
    Exit Sub
ErrorHandle:
    custErr.RaiseError "cTable - ClearTableContents()"
End Sub

Sub DeleteTableRows()
    ' Description: Deletes all rows in the table
    ' Dependencies:
    '   - Me.lo()
    ' Inputs: None
    ' Outputs: None
    On Error GoTo ErrorHandle
    If Me.NumberRows > 0 Then
        Me.lo.AutoFilter.ShowAllData
        Me.lo.DataBodyRange.Delete
    End If
    Exit Sub
ErrorHandle:
    custErr.RaiseError "cTable - DeleteTableRows()"
End Sub

' Private Get Properties
Private Property Get pvtHeaderCell(ByVal ColName As String) As Range
    ' Description: Returns a cell in the header matching the given column name (case sensitive)
    ' Dependencies:
    '   - Me.lo()
    ' Inputs:
    '   - (As String)   ColName     - The name of the column to identify
    ' Outputs:
    '   - (As Range) The cell in the header matching the column name given
    On Error GoTo ErrorHandle
    Set pvtHeaderCell = Me.lo.HeaderRowRange.Find(what:=ColName, LookIn:=XlFindLookIn.xlValues, LookAt:=XlLookAt.xlWhole, MatchCase:=True)
    Exit Property
ErrorHandle:
    custErr.RaiseError "cTable - Get pvtHeaderCell()"
End Property

' Public Set Properties
Property Set lo(ByRef lo As ListObject)
    ' Description: Defines the table as an object
    ' Dependencies: None
    ' Inputs:
    '   - (As ListObject)   lo  - The table object
    ' Outputs: None
    On Error GoTo ErrorHandle
    Set vTbl = lo
    vShowFilter = vTbl.ShowAutoFilter
    vExistingRows = vTbl.ListRows.Count
    Exit Property
ErrorHandle:
    custErr.RaiseError "cTable - Set lo()"
End Property

' Public Get Properties
Property Get lo() As ListObject
    ' Description: Returns the table as a listobject
    ' Dependencies:
    '   - Table must have been defined
    ' Inputs: None
    ' Outputs:
    '   - (As ListObject) The defined table
    On Error GoTo ErrorHandle
    If vTbl Is Nothing Then Err.Raise custErr.GenericError, Description:="Table has not been defined"
    Set lo = vTbl
    Exit Property
ErrorHandle:
    custErr.RaiseError "cTable - Get lo()"
End Property

Property Get Sheet() As Worksheet
    ' Description: Returns the worksheet that the table is in
    ' Dependencies:
    '   - Me.lo()
    ' Inputs: None
    ' Outputs:
    '   - (As Worksheet) The worksheet that the table is in
    On Error GoTo ErrorHandle
    Set Sheet = Me.lo.Parent
    Exit Property
ErrorHandle:
    custErr.RaiseError "cTable - Get Sheet()"
End Property

Property Get wb() As Workbook
    ' Description: Returns the workbook that the table is in
    ' Dependencies:
    '   - Me.Sheet()
    ' Inputs: None
    ' Outputs:
    '   - (As Workbook) The workbook that the table is in
    On Error GoTo ErrorHandle
    Set wb = Me.Sheet.Parent
    Exit Property
ErrorHandle:
    custErr.RaiseError "cTable - Get wb()"
End Property

Property Get HeaderRow() As Long
    ' Description: The sheet row that the table's header is in
    ' Dependencies:
    '   - Me.lo()
    ' Inputs: None
    ' Outputs:
    '   - (As Long) The sheet row number that the header is in
    On Error GoTo ErrorHandle
    Me.lo.HeaderRowRange.Row
    Exit Property
ErrorHandle:
    custErr.RaiseError "cTable - Get HeaderRow()"
End Property

Property Get NumberColumns() As Long
    ' Description: The number of columns in the table
    ' Dependencies:
    '   - Me.lo()
    ' Inputs: None
    ' Outputs:
    '   - (As Long) The number of columns in the table
    On Error GoTo ErrorHandle
    NumberColumns = Me.lo.ListColumns.Count
    Exit Property
ErrorHandle:
    custErr.RaiseError "cTable - Get NumberColumns()"
End Property

Property Get NumberRows() As Long
    ' Description: The number of rows of data in the table
    ' Dependencies:
    '   - Me.lo()
    ' Inputs: None
    ' Outputs:
    '   - (As Long) The number of rows of data in the table
    On Error GoTo ErrorHandle
    NumberRows = Me.lo.ListRows.Count
    Exit Property
ErrorHandle:
    custErr.RaiseError "cTable - Get NumberRows()"
End Property

Property Get FirstDataRow() As Long
    ' Description: The sheet row number of the first row of data in the table
    ' Dependencies:
    '   - Me.lo()
    '   - Table must contain data
    ' Inputs: None
    ' Outputs:
    '   - (As Long) The sheet row number of the first row of data in the table
    On Error GoTo ErrorHandle
    FirstDataRow = Me.lo.ListRows(1).Range.Row
    Exit Property
ErrorHandle:
    custErr.RaiseError "cTable - Get FirstDataRow()"
End Property

Property Get LastDataRow() As Long
    ' Description: The sheet row number of the last row of data in the table
    ' Dependencies:
    '   - Me.lo()
    '   - Table must contain data
    ' Inputs: None
    ' Outputs:
    '   - (As Long) The sheet row number of the last row of data in the table
    On Error GoTo ErrorHandle
    FirstDataRow = Me.lo.ListRows(Me.lo.ListRows.Count).Range.Row
    Exit Property
ErrorHandle:
    custErr.RaiseError "cTable - Get LastDataRow()"
End Property

Property Get FirstColumnNum() As Long
    ' Description: The sheet column number of the first column in the table
    ' Dependencies:
    '   - Me.lo()
    ' Inputs: None
    ' Outputs:
    '   - (As Long) The sheet column number of the first column in the table
    On Error GoTo ErrorHandle
    FirstColumnNum = Me.lo.ListColumns(1).Range.Column
    Exit Property
ErrorHandle:
    custErr.RaiseError "cTable - Get FirstColumnNum()"
End Property

Property Get LastColumnNum() As Long
    ' Description: The sheet column number of the last column in the table
    ' Dependencies:
    '   - Me.lo()
    ' Inputs: None
    ' Outputs:
    '   - (As Long) The sheet column number of the last column in the table
    On Error GoTo ErrorHandle
    FirstColumnNum = Me.lo.ListColumns(Me.lo.ListColumns.Count).Range.Column
    Exit Property
ErrorHandle:
    custErr.RaiseError "cTable - Get LastColumnNum()"
End Property

Property Get FirstColumnLetter() As String
    ' Description: The sheet column letter of the first column in the table
    ' Dependencies:
    '   - Me.FirstColumnNum()
    ' Inputs: None
    ' Outputs:
    '   - (As String) The sheet column letter of the first column in the table
    On Error GoTo ErrorHandle
    FirstColumnLetter = Replace(Cells(1, Me.FirstColumnNum).Address(RowAbsolute:=False, ColumnAbsolute:=False), 1, "")
    Exit Property
ErrorHandle:
    custErr.RaiseError "cTable - Get FirstColumnLetter()"
End Property

Property Get LastColumnLetter() As String
    ' Description: The sheet column letter of the last column in the table
    ' Dependencies:
    '   - Me.LastColumnNum()
    ' Inputs: None
    ' Outputs:
    '   - (As String) The sheet column letter of the last column in the table
    On Error GoTo ErrorHandle
    LastColumnLetter = Replace(Cells(1, Me.LastColumnNum).Address(RowAbsolute:=False, ColumnAbsolute:=False), 1, "")
    Exit Property
ErrorHandle:
    custErr.RaiseError "cTable - Get LastColumnLetter()"
End Property

Property Get ColumnPresent(ByVal ColName As String) As Boolean
    ' Description: Returns whether a column with a given title exists in the table or not
    ' Dependencies:
    '   - private pvtHeaderCell()
    ' Inputs:
    '   - (As String)   ColName     - The name of the column to identify (case sensitive)
    ' Outputs:
    '   - (As Boolean) Whether the named column exists or not
    On Error GoTo ErrorHandle
    ColumnPresent = pvtHeaderCell(ColName) Is Not Nothing
    Exit Property
ErrorHandle:
    custErr.RaiseError "cTable - Get ColumnPresent()"
End Property

Property Get HeaderCell(ByVal ColName As String) As Range
    ' Description: Returns the cell in the header with the given column name
    ' Dependencies:
    '   - private pvtHeaderCell()
    '   - Me.ColumnPresent()
    ' Inputs:
    '   - (As String)   ColName     - The name of the column to identify (case sensitive)
    ' Outputs:
    '   - (As Range) The cell in the header with the given column name
    On Error GoTo ErrorHandle
    If Not Me.ColumnPresent(ColName) Then Err.Raise custErr.GenericError, Description:="The column, " & ColName & ", has not been found"
    Set HeaderCell = pvtHeaderCell(ColName)
    Exit Property
ErrorHandle:
    custErr.RaiseError "cTable - Get HeaderCell()"
End Property

Property Get ColumnDataColNum(ByVal ColName As String) As Long
    ' Description: Returns the data column number of the given column name
    ' Dependencies:
    '   - Me.ColumnSheetColNum()
    '   - Me.FirstColumnNum()
    ' Inputs:
    '   - (As String)   ColName     - The name of the column to identify (case sensitive)
    ' Outputs:
    '   - (As Long) The data column number with the given column name
    On Error GoTo ErrorHandle
    ColumnDataColNum = Me.ColumnSheetColNum(ColName) - Me.FirstColumnNum + 1
    Exit Property
ErrorHandle:
    custErr.RaiseError "cTable - Get ColumnDataColNum()"
End Property

Property Get ColumnSheetColNum(ByVal ColName As String) As Long
    ' Description: Returns the sheet column number of the given column name
    ' Dependencies:
    '   - Me.HeaderCell()
    ' Inputs:
    '   - (As String)   ColName     - The name of the column to identify (case sensitive)
    ' Outputs:
    '   - (As Long) The sheet column number with the given column name
    On Error GoTo ErrorHandle
    ColumnSheetColNum = Me.HeaderCell(ColName).Column
    Exit Property
ErrorHandle:
    custErr.RaiseError "cTable - Get ColumnSheetColNum()"
End Property

Property Get ColumnSheetColLetter(ByVal ColName As String) As String
    ' Description: Returns the sheet column letter of the given column name
    ' Dependencies:
    '   - Me.ColumnSheetColNum()
    ' Inputs:
    '   - (As String)   ColName     - The name of the column to identify (case sensitive)
    ' Outputs:
    '   - (As String) The sheet column letter with the given column name
    On Error GoTo ErrorHandle
    ColumnSheetColLetter = Replace(Cells(1, Me.ColumnSheetColNum(ColName)).Address(RowAbsolute:=False, ColumnAbsolute:=False), 1, "")
    Exit Property
ErrorHandle:
    custErr.RaiseError "cTable - Get ColumnSheetColLetter()"
End Property

Property Get ColumnDataCol(ByVal ColName As String) As ListColumn
    ' Description: Returns the column of data with the given column name
    ' Dependencies:
    '   - Me.lo()
    '   - Me.ColumnPresent()
    ' Inputs:
    '   - (As String)   ColName     - The name of the column to identify (case sensitive)
    ' Outputs:
    '   - (As ListColumn) The table column with the given header name
    On Error GoTo ErrorHandle
    If Not Me.ColumnPresent(ColName) Then Err.Raise custErr.GenericError, Description:="The column, " & ColName & ", has not been found"
    Set ColumnDataCol = Me.lo.ListColumns(ColName)
    Exit Property
ErrorHandle:
    custErr.RaiseError "cTable - Get ColumnDataCol()"
End Property

Property Get ColumnWholeSheetCol(ByVal ColName As String) As Range
    ' Description: Returns the whole sheet column with the given column name
    ' Dependencies:
    '   - Me.HeaderCell()
    ' Inputs:
    '   - (As String)   ColName     - The name of the column to identify (case sensitive)
    ' Outputs:
    '   - (As Range) The whole sheet column with the given column name
    On Error GoTo ErrorHandle
    Me.HeaderCell(ColName).EntireColumn
    Exit Property
ErrorHandle:
    custErr.RaiseError "cTable - Get ColumnWholeSheetCol()"
End Property
