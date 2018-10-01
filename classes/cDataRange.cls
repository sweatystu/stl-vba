Option Explicit

' Dependencies:
'   - cCustomErrors

' Private Variables
Private vDRng As Range

' Public Set Properties
Property Set DataRange(ByRef rng As Range)
    ' Description: Define the data range - Range underneath the header
    ' Dependencies: None
    ' Inputs:
    '   - (As Range)    rng - The range of cells to be recorded
    ' Outputs: None
    On Error GoTo ErrorHandle
    Set vDRng = rng
    Exit Property
ErrorHandle:
    custErr.RaiseError "cDataRange - Set DataRange()"
End Property

' Public Get Properties
Property Get DataRange() As Range
    ' Description: Return the data range
    ' Dependencies:
    '   - The data range must have been defined
    ' Inputs: None
    ' Outputs:
    '   - (As Range) The range of cells defined
    On Error GoTo ErrorHandle
    If vDRng Is Nothing Then Err.Raise custErr.GenericError, Description:="The range has not been defined"
    Set DataRange = vDRng
    Exit Property
ErrorHandle:
    custErr.RaiseError "cDataRange - Get DataRange()"
End Property

Property Get NumberRows() As Long
    ' Description: Returns the number of rows in the data range
    ' Dependencies:
    '   - Me.DataRange()
    ' Inputs: None
    ' Outputs:
    '   - (As Long) The number of rows in the range
    On Error GoTo ErrorHandle
    NumberRows = Me.DataRange.Rows.Count
    Exit Property
ErrorHandle:
    custErr.RaiseError "cDataRange - Get NumberRows()"
End Property


Property Get FirstRow() As Long
    ' Description: Returns the sheet number of the first row in the data range
    ' Dependencies:
    '   - Me.DataRange()
    ' Inputs: None
    ' Outputs:
    '   - (As Long) The sheet number of the first row in the range
    On Error GoTo ErrorHandle
    FirstRow = Me.DataRange.Rows(1).Row
    Exit Property
ErrorHandle:
    custErr.RaiseError "cDataRange - Get FirstRow()"
End Property

Property Get LastRow() As Long
    ' Description: Returns the sheet number of the last row in the data range
    ' Dependencies:
    '   - Me.DataRange()
    ' Inputs: None
    ' Outputs:
    '   - (As Long) The sheet number of the last row in the range
    On Error GoTo ErrorHandle
    FirstRow = Me.DataRange.Rows(Me.NumberRows).Row
    Exit Property
ErrorHandle:
    custErr.RaiseError "cDataRange - Get FirstRow()"
End Property


Property Get RowInData(ByVal DataRow As Long) As Range
    ' Description: Returns the nth row of data in the range as an object
    ' Dependencies:
    '   - Me.NumberRows()
    '   - Me.DataRange()
    ' Inputs:
    '   - (As Long) DataRow - must be between 1 and the number of rows in the data
    ' Outputs:
    '   - (As Range) The nth row of data
    On Error GoTo ErrorHandle
    If DataRow < 1 Then Err.Raise custErr.GenericError, Description:="The number row to return must be greater than 0"
    If DataRow > Me.NumberRows Then Err.Raise custErr.GenericError, Description:="There are not enough rows in the data (" & Me.NumberRows & ") for the value given (" & DataRow & ")"
    Set RowInData = Me.DataRange.Rows(DataRow)
    Exit Property
ErrorHandle:
    custErr.RaiseError "cDataRange - Get RowInData()"
End Property

Property Get RowInSheet(ByVal SheetRow As Long) As Range
    ' Description: Returns the nth sheet row of data in the range as an object
    ' Dependencies:
    '   - Me.DataRange()
    ' Inputs:
    '   - (As Long) SheetRow    - must intersect with the data range
    ' Outputs:
    '   - (As Range) The nth sheet row of data (intersect)
    Dim result As Range
    On Error GoTo ErrorHandle
    Set result = Intersect(Me.DataRange, Me.DataRange.Parent.Rows(SheetRow))
    If result Is Nothing Then Err.Raise custErr.GenericError, Description:="The sheet row given (" & SheetRow & ") is not within the data range (" & Me.FirstRow & " - " & Me.LastRow & ")"
    Set RowInSheet = result
    Exit Property
ErrorHandle:
    custErr.RaiseError "cDataRange - Get RowInSheet()"
End Property

Property Get ColumnInData(ByVal dataCol As Long) As Range
    ' Description: Returns the nth column of data in the range as an object
    ' Dependencies:
    '   - Me.DataRange()
    ' Inputs:
    '   - (As Long) dataCol - must be between 1 and the number of columns in the data
    ' Outputs:
    '   - (As Range) The nth column of data
    On Error GoTo ErrorHandle
    If dataCol < 1 Then Err.Raise custErr.GenericError, Description:="The column number must be greater than 0"
    If dataCol > Me.DataRange.Columns.Count Then Err.Raise custErr.GenericError, Description:="There are not enough columns in the data (" & vDRng.Columns.Count & ") for the value given (" & dataCol & ")"
    Set ColumnInData = Me.DataRange.Columns(dataCol)
    Exit Property
ErrorHandle:
    custErr.RaiseError "cDataRange - Get ColumnInData()"
End Property

Property Get ColumnInSheetNum(ByVal SheetCol As Long) As Range
    ' Description: Returns the nth sheet column of data in the range as an object
    ' Dependencies:
    '   - Me.DataRange()
    ' Inputs:
    '   - (As Long) SheetCol    - must intersect with the data range
    ' Outputs:
    '   - (As Range) The nth sheet column of data (intersect)
    Dim result As Range
    On Error GoTo ErrorHandle
    Set result = Intersect(Me.DataRange, Me.DataRange.Parent.Columns(SheetCol))
    If result Is Nothing Then Err.Raise custErr.GenericError, Description:="The sheet column number given (" & SheetCol & ") does not intersect with the data range"
    Set ColumnInSheetNum = result
    Exit Property
ErrorHandle:
    custErr.RaiseError "cDataRange - Get ColumnInSheetNum()"
End Property

Property Get ColumnInSheetLetter(ByVal SheetCol As String) As Range
    ' Description: Returns the nth sheet column (as a letter) of data in the range as an object
    ' Dependencies:
    '   - Me.DataRange()
    ' Inputs:
    '   - (As String) SheetCol  - Letter of the sheet column
    '                           - must intersect with the data range
    ' Outputs:
    '   - (As Range) The nth sheet column (as a letter) of data (intersect)
    Dim tempcl As Range
    Dim result As Range
    On Error GoTo ErrorHandle
    Set tempcl = Range(SheetCol & 1)
    If tempcl Is Nothing Then Err.Raise custErr.GenericError, Description:="The sheet column letter given (" & SheetCol & ") has not been recognised as a column"
    Set result = Intersect(Me.DataRange, Me.DataRange.Parent.Columns(tempcl.Column))
    If result Is Nothing Then Err.Raise custErr.GenericError, Description:="The sheet column letter given (" & SheetCol & ") does not intersect with the data range"
    Set ColumnInSheetNum = result
    Exit Property
ErrorHandle:
    custErr.RaiseError "cDataRange - Get ColumnInSheetLetter()"
End Property

Property Get Cell(ByVal RowNum As Long, ByVal ColNum As Long) As Range
    ' Description: Returns a cell in the range defined by the row number and the column number
    ' Dependencies:
    '   - Me.DataRange()
    ' Inputs:
    '   - (As Long) RowNum  - The numbered row in the data to return a result from
    '   - (As Long) ColNum  - The numbered column in the data to return a result from
    ' Outputs:
    '   - (As Range) The cell in the nth row and nth column of data
    On Error GoTo ErrorHandle
    If RowNum < 1 Or ColNum < 1 Then Err.Raise custErr.GenericError, Description:="The row number (" & RowNum & ") and column number (" & ColNum & ") must be greater than 0"
    If RowNum > Me.NumberRows Or ColNum > Me.DataRange.Columns.Count Then Err.Raise custErr.GenericError, Description:="The row and column numbers given (" & RowNum & " & " & ColNum & ") are beyond the scope of the range (" & Me.NumberRows & " rows, " & Me.DataRange.Columns.Count & " columns)"
    Set Cell = Me.DataRange.Cells(RowNum, ColNum)
    Exit Property
ErrorHandle:
    custErr.RaiseError "cDataRange - Get Cell()"
End Property
