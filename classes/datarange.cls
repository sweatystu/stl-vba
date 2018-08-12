Option Explicit

Private vDRng As Range


Property Set DataRange(ByRef rng As Range)
    On Error GoTo ErrorHandle
    Set vDRng = rng
    Exit Property
ErrorHandle:
    custErr.RaiseError "cDataRange - Set DataRange()"
End Property

Property Get DataRange() As Range
    On Error GoTo ErrorHandle
    If vDRng Is Nothing Then Err.Raise custErr.GenericError, Description:="The range has not been defined"
    Set DataRange = vDRng
    Exit Property
ErrorHandle:
    custErr.RaiseError "cDataRange - Get DataRange()"
End Property


Property Get NumberRows() As Long
    On Error GoTo ErrorHandle
    NumberRows = Me.DataRange.Rows.Count
    Exit Property
ErrorHandle:
    custErr.RaiseError "cDataRange - Get NumberRows()"
End Property


Property Get FirstRow() As Long
    On Error GoTo ErrorHandle
    FirstRow = Me.DataRange.Rows(1).row
    Exit Property
ErrorHandle:
    custErr.RaiseError "cDataRange - Get FirstRow()"
End Property

Property Get LastRow() As Long
    On Error GoTo ErrorHandle
    FirstRow = Me.DataRange.Rows(Me.NumberRows).row
    Exit Property
ErrorHandle:
    custErr.RaiseError "cDataRange - Get FirstRow()"
End Property


Property Get RowInData(ByVal DataRow As Long) As Range
    On Error GoTo ErrorHandle
    If DataRow > Me.NumberRows Then Err.Raise custErr.GenericError, Description:="There are not enough rows in the data (" & Me.NumberRows & ") for the value given (" & DataRow & ")"
    Set RowInData = Me.DataRange.Rows(DataRow)
    Exit Property
ErrorHandle:
    custErr.RaiseError "cDataRange - Get RowInData()"
End Property

Property Get RowInSheet(ByVal SheetRow As Long) As Range
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
    On Error GoTo ErrorHandle
    If dataCol > Me.DataRange.Columns.Count Then Err.Raise custErr.GenericError, Description:="There are not enough columns in the data (" & vDRng.Columns.Count & ") for the value given (" & dataCol & ")"
    Set ColumnInData = Me.DataRange.Columns(dataCol)
    Exit Property
ErrorHandle:
    custErr.RaiseError "cDataRange - Get ColumnInData()"
End Property

Property Get ColumnInSheetNum(ByVal SheetCol As Long) As Range
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
    On Error GoTo ErrorHandle
    If RowNum < 1 Or ColNum < 1 Then Err.Raise custErr.GenericError, Description:="The row number (" & RowNum & ") and column number (" & ColNum & ") must be greater than 0"
    If RowNum > Me.NumberRows Or ColNum > Me.DataRange.Columns.Count Then Err.Raise custErr.GenericError, Description:="The row and column numbers given (" & RowNum & " & " & ColNum & ") are beyond the scope of the range (" & Me.NumberRows & " rows, " & Me.DataRange.Columns.Count & " columns)"
    Set Cell = Me.DataRange.Cells(RowNum, ColNum)
    Exit Property
ErrorHandle:
    custErr.RaiseError "cDataRange - Get Cell()"
End Property
