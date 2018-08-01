Option Explicit

Private vRng As Range
Private vHeader As Boolean
Private vHeaderRow As Long



Sub DefineRange(ByRef CellRange As Range, Optional HasHeaderRow As Boolean = True, Optional HeaderRow As Long = 1)
    On Error GoTo ErrorHandle
    
    Set vRng = CellRange
    vHeader = HasHeaderRow
    If vHeader Then
10      If vHeader < 1 Then Err.Raise custErr.GenericError, Description:="Range header-row number is not valid"
20      If vHeader > vRng.Rows.Count Then Err.Raise custErr.GenericError, Description:="Range header-row number is not included within the range"
        vHeaderRow = HeaderRow
    Else
        vHeaderRow = 0
    End If
    
    Exit Sub
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

Property Get sheet() As Worksheet
    On Error GoTo ErrorHandle
    If vRng Is Nothing Then Err.Raise custErr.GenericError, Description:="The range has not been defined"
    Set sheet = vRng.Parent
    Exit Property
ErrorHandle:
    custErr.RaiseError "cRange - Get sheet()"
End Property

Property Get wb() As Workbook
    On Error GoTo ErrorHandle
    If vRng Is Nothing Then Err.Raise custErr.GenericError, Description:="The range has not been defined"
    Set wb = Me.sheet.Parent
    Exit Property
ErrorHandle:
    custErr.RaiseError "cRange - Get sheet()"
End Property


' Sub-Ranges

Property Get HeaderRange() As Range
    On Error GoTo ErrorHandle
    If vRng Is Nothing Then Err.Raise custErr.GenericError, Description:="The range has not been defined"
    If Not vHeader Then Err.Raise custErr.GenericError, Description:="The range does not have a header"
    Set HeaderRange = vRng.Rows(vHeaderRow)
    Exit Property
ErrorHandle:
    custErr.RaiseError "cRange - Get HeaderRange()"
End Property

Property Get DataRange() As Range
    Dim rng As Range
    On Error GoTo ErrorHandle
    If vRng Is Nothing Then Err.Raise custErr.GenericError, Description:="The range has not been defined"
    If Not vHeader Then
        Set rng = vRng
    Else
        If Me.HeaderRow >= Me.LastRow Then Err.Raise custErr.GenericError, Description:="There is no data in the range"
        Set rng = Intersect(vRng, Me.sheet.Rows(Me.FirstRow & ":" & Me.LastRow))
    End If
    Set DataRange = rng
    Exit Property
ErrorHandle:
    custErr.RaiseError "cRange - Get DataRange()"
End Property

Property Get ColumnDataRange(ByVal ColName As String) As Range
    On Error GoTo ErrorHandle
    If vRng Is Nothing Then Err.Raise custErr.GenericError, Description:="The range has not been defined"
    If Me.HeaderRow >= Me.LastRow Then Err.Raise custErr.GenericError, Description:="There is no data in the range"
    If Not ColumnPresent(ColName) Then Err.Raise custErr.GenericError, Description:="The column " & ColName & " could not be found"
    Set ColumnDataRange = Intersect(Me.DataRange, Me.HeaderRange.Find(What:=ColName, LookIn:=XlFindLookIn.xlValues, LookAt:=XlLookAt.xlWhole, MatchCase:=True).EntireColumn)
    Exit Property
ErrorHandle:
    custErr.RaiseError "cRange - Get ColumnDataRange()"
End Property

Property Get ColumnHeader(ByVal ColName As String) As Range
    On Error GoTo ErrorHandle
    If vRng Is Nothing Then Err.Raise custErr.GenericError, Description:="The range has not been defined"
    If Not vHeader Then Err.Raise custErr.GenericError, Description:="The range does not have a header"
    If Not ColumnPresent(ColName) Then Err.Raise custErr.GenericError, Description:="The column " & ColName & " could not be found"
    Set ColumnHeader = Me.HeaderRange.Find(What:=ColName, LookIn:=XlFindLookIn.xlValues, LookAt:=XlLookAt.xlWhole, MatchCase:=True)
    Exit Property
ErrorHandle:
    custErr.RaiseError "cRange - Get ColumnHeader()"
End Property



' Properties

Property Get HeaderRow() As Long
    On Error GoTo ErrorHandle
    If vRng Is Nothing Then Err.Raise custErr.GenericError, Description:="The range has not been defined"
    HeaderSheetRow = vRng.Rows(vHeaderRow).row
    Exit Property
ErrorHandle:
    custErr.RaiseError "cRange - Get HeaderRow()"
End Property


Property Get FirstRow() As Long
    Dim row As Long
    On Error GoTo ErrorHandle
    If vRng Is Nothing Then Err.Raise custErr.GenericError, Description:="The range has not been defined"
    If vHeader Then
        If Me.HeaderRow >= Me.LastRow Then Err.Raise custErr.GenericError, Description:="There is no data in the range"
        row = Me.HeaderRow + 1
    Else
        row = vRng.Rows(1).row
    End If
    FirstRow = row
    Exit Property
ErrorHandle:
    custErr.RaiseError "cRange - Get FirstRow()"
End Property

Property Get LastRow() As Long
    On Error GoTo ErrorHandle
    If vRng Is Nothing Then Err.Raise custErr.GenericError, Description:="The range has not been defined"
    LastRow = vRng.Rows(vRng.Rows.Count).row
    Exit Property
ErrorHandle:
    custErr.RaiseError "cRange - Get LastRow()"
End Property

Property Get FirstCol() As Long
    On Error GoTo ErrorHandle
    If vRng Is Nothing Then Err.Raise custErr.GenericError, Description:="The range has not been defined"
    FirstCol = vRng.Columns(1).Column
    Exit Property
ErrorHandle:
    custErr.RaiseError "cRange - Get FirstCol()"
End Property

Property Get LastCol() As Long
    On Error GoTo ErrorHandle
    If vRng Is Nothing Then Err.Raise custErr.GenericError, Description:="The range has not been defined"
    LastCol = vRng.Columns(vRng.Columns.Count).Column
    Exit Property
ErrorHandle:
    custErr.RaiseError "cRange - Get LastCol()"
End Property

Property Get FirstColLetter() As String
    On Error GoTo ErrorHandle
    FirstColLetter = Replace(Cells(1, Me.FirstCol).Address(RowAbsolute:=False, ColumnAbsolute:=False), 1, "")
    Exit Property
ErrorHandle:
    custErr.RaiseError "cRange - Get FirstColLetter()"
End Property

Property Get LastColLetter() As String
    On Error GoTo ErrorHandle
    LastColLetter = Replace(Cells(1, Me.LastCol).Address(RowAbsolute:=False, ColumnAbsolute:=False), 1, "")
    Exit Property
ErrorHandle:
    custErr.RaiseError "cRange - Get LastColLetter()"
End Property


' Columns
Private Function ColumnPresent(ByVal ColName As String) As Boolean
    On Error GoTo ErrorHandle
    If vRng Is Nothing Then Err.Raise custErr.GenericError, Description:="The range has not been defined"
    If Not vHeader Then Err.Raise custErr.GenericError, Description:="The range does not have a header"
    ColumnPresent = Not Me.HeaderRange.Find(What:=ColName, LookIn:=XlFindLookIn.xlValues, LookAt:=XlLookAt.xlWhole, MatchCase:=True) Is Nothing
    Exit Function
ErrorHandle:
    custErr.RaiseError "cRange - ColumnPresent()"
End Function

Property Get ColumnNumber(ByVal ColName As String) As Long
    On Error GoTo ErrorHandle
    If vRng Is Nothing Then Err.Raise custErr.GenericError, Description:="The range has not been defined"
    If Not vHeader Then Err.Raise custErr.GenericError, Description:="The range does not have a header"
    If Not ColumnPresent(ColName) Then Err.Raise custErr.GenericError, Description:="The column " & ColName & " could not be found"
    ColumnNumber = Me.HeaderRange.Find(What:=ColName, LookIn:=XlFindLookIn.xlValues, LookAt:=XlLookAt.xlWhole, MatchCase:=True).Column
    Exit Property
ErrorHandle:
    custErr.RaiseError "cRange - Get ColumnNumber()"
End Property

Property Get ColumnLetter(ByVal ColName As String) As String
    On Error GoTo ErrorHandle
    If vRng Is Nothing Then Err.Raise custErr.GenericError, Description:="The range has not been defined"
    If Not vHeader Then Err.Raise custErr.GenericError, Description:="The range does not have a header"
    If Not ColumnPresent(ColName) Then Err.Raise custErr.GenericError, Description:="The column " & ColName & " could not be found"
    ColumnLetter = Replace(Cells(1, Me.ColumnNumber(ColName)).Address(RowAbsolute:=False, ColumnAbsolute:=False), 1, "")
    Exit Property
ErrorHandle:
    custErr.RaiseError "cRange - Get ColumnLetter()"
End Property

