Option Explicit

Private vRng As Range
Private vHeaderRange As cHeaderRange
Private vDataRange As cDataRange


' Definition

Sub DefineRange(ByRef CellRange As Range, Optional HasHeaderRow As Boolean = True, Optional HeaderRow As Long = 1)
    On Error GoTo ErrorHandle
    Set vRng = CellRange
    If HasHeaderRow Then
        If HeaderRow < 1 Then Err.Raise custErr.GenericError, Description:="The header-row number (" & HeaderRow & ") must be greater than 0"
        If HeaderRow >= vRng.Rows.Count Then Err.Raise custErr.GenericError, Description:="The header-row number (" & HeaderRow & ") must be less than the number of rows in the data (" & vRng.Rows.Count & ")"
        Set vHeaderRange = New cHeaderRange
        Set vHeaderRange.HeaderRange = vRng.Rows(HeaderRow)
        Set vDataRange = New cDataRange
        Set vDataRange.DataRange = Intersect(vRng, vRng.Parent.Rows(vHeaderRange.SheetRow + 1 & ":" & vRng.Rows(vRng.Rows.Count).row))
    Else
        Set vHeaderRange = Nothing
        Set vDataRange = vRng
    End If
ErrorHandle:
    custErr.RaiseError "cRange - DefineRange()"
End Sub


' Objects

Property Get rng() As Range
    On Error GoTo ErrorHandle
    If vRng Is Nothing Then Err.Raise custErr.GenericError, Description:="The range has not been defined"
    Set rng = vRng
    Exit Property
ErrorHandle:
    custErr.RaiseError "cRange - Get rng()"
End Property

Property Get Sheet() As Worksheet
    On Error GoTo ErrorHandle
    If vRng Is Nothing Then Err.Raise custErr.GenericError, Description:="The range has not been defined"
    Set Sheet = vRng.Parent
    Exit Property
ErrorHandle:
    custErr.RaiseError "cRange - Get sheet()"
End Property

Property Get wb() As Workbook
    On Error GoTo ErrorHandle
    If vRng Is Nothing Then Err.Raise custErr.GenericError, Description:="The range has not been defined"
    Set wb = Me.Sheet.Parent
    Exit Property
ErrorHandle:
    custErr.RaiseError "cRange - Get sheet()"
End Property

Property Get HeaderRange() As cHeaderRange
    On Error GoTo ErrorHandle
    If vRng Is Nothing Then Err.Raise custErr.GenericError, Description:="The range has not been defined"
    If vHeaderRange Is Nothing Then Err.Raise custErr.GenericError, Description:="The range does not have a header"
    Set HeaderRange = vHeaderRange
    Exit Property
ErrorHandle:
    custErr.RaiseError "cRange - Get HeaderRange()"
End Property

Property Get DataRange() As cDataRange
    Dim rng As Range
    On Error GoTo ErrorHandle
    If vRng Is Nothing Then Err.Raise custErr.GenericError, Description:="The range has not been defined"
    If vDataRange Is Nothing Then Err.Raise custErr.GenericError, Description:="The data range has not been defined"
    Set DataRange = vDataRange
    Exit Property
ErrorHandle:
    custErr.RaiseError "cRange - Get DataRange()"
End Property



' Column Properties

Property Get NoColumns() As Long
    On Error GoTo ErrorHandle
    If vRng Is Nothing Then Err.Raise custErr.GenericError, Description:="The range has not been defined"
    NoColumns = vRng.Columns.Count
    Exit Property
ErrorHandle:
    custErr.RaiseError "cRange - Get NoColumns()"
End Property

Property Get FirstColNum() As Long
    On Error GoTo ErrorHandle
    If vRng Is Nothing Then Err.Raise custErr.GenericError, Description:="The range has not been defined"
    FirstCol = vRng.Columns(1).Column
    Exit Property
ErrorHandle:
    custErr.RaiseError "cRange - Get FirstCol()"
End Property

Property Get LastColNum() As Long
    On Error GoTo ErrorHandle
    If vRng Is Nothing Then Err.Raise custErr.GenericError, Description:="The range has not been defined"
    LastCol = vRng.Columns(vRng.Columns.Count).Column
    Exit Property
ErrorHandle:
    custErr.RaiseError "cRange - Get LastCol()"
End Property

Property Get FirstColLetter() As String
    On Error GoTo ErrorHandle
    If vRng Is Nothing Then Err.Raise custErr.GenericError, Description:="The range has not been defined"
    FirstColLetter = Replace(Cells(1, Me.FirstColNum).Address(RowAbsolute:=False, ColumnAbsolute:=False), 1, "")
    Exit Property
ErrorHandle:
    custErr.RaiseError "cRange - Get FirstColLetter()"
End Property

Property Get LastColLetter() As String
    On Error GoTo ErrorHandle
    If vRng Is Nothing Then Err.Raise custErr.GenericError, Description:="The range has not been defined"
    LastColLetter = Replace(Cells(1, Me.LastColNum).Address(RowAbsolute:=False, ColumnAbsolute:=False), 1, "")
    Exit Property
ErrorHandle:
    custErr.RaiseError "cRange - Get LastColLetter()"
End Property



' Cells

Property Get Cell_DataRowColName(ByVal DataRow As Long, ByVal ColName As Long) As Range
    Dim result As Range
    On Error GoTo ErrorHandle
    Set result = Intersect(Me.DataRange.RowInData(DataRow), Me.HeaderRange.WholeSheetCol(ColName))
    If result Is Nothing Then Err.Raise custErr.GenericError, Description:="The row (" & DataRow & ") and column (" & ColName & ") combination has not been recognised"
    Set Cell_DataRowColName = result
    Exit Property
ErrorHandle:
    custErr.RaiseError "cRange - Get Cell_DataRowColName()"
End Property

Property Get Cell_SheetRowColName(ByVal SheetRow As Long, ByVal ColName As Long) As Range
    Dim result As Range
    On Error GoTo ErrorHandle
    Set result = Intersect(Me.Sheet.Rows(SheetRow), Me.HeaderRange.WholeSheetCol(ColName))
    If result Is Nothing Then Err.Raise custErr.GenericError, Description:="The sheet row (" & SheetRow & ") and column (" & ColName & ") combination has not been recognised"
    Set Cell_SheetRowColName = result
    Exit Property
ErrorHandle:
    custErr.RaiseError "cRange - Get Cell_SheetRowColName()"
End Property

Property Get Cell_UnderHeader(ByVal RowNum As Long, ByVal ColName As String) As Range
    Dim result As Range
    On Error GoTo ErrorHandle
    If RowNum < 1 Then Err.Raise custErr.GenericError, Description:="The row number given (" & RowNum & ") should be greater than 0"
    Set result = Me.HeaderRange.ColCell(ColName).Offset(RowNum, 0)
    If result Is Nothing Then Err.Raise custErr.GenericError, Description:="The sheet row (" & SheetRow & ") and column (" & ColName & ") combination has not been recognised"
    Set Cell_UnderHeader = result
    Exit Property
ErrorHandle:
    custErr.RaiseError "cRange - Get Cell_UnderHeader()"
End Property


