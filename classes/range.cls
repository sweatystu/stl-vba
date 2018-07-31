Option Explicit

Private vRng As Range
Private vHeader As Boolean
Private vHeaderRow As Long



Sub DefineRange(ByRef CellRange As Range, Optional HasHeaderRow As Boolean = True, Optional HeaderRow As Long = 1)
    On Error GoTo ErrorHandle
    
    Set vRng = CellRange
    vHeader = HasHeaderRow
    If vHeader Then
10        If vHeader < 1 Then Err.Raise custErr.GenericError, Description:="Range header-row number is not valid"
20        If vHeader > vRng.Rows.Count Then Err.Raise custErr.GenericError, Description:="Range header-row number is not included within the range"
        vHeaderRow = HeaderRow
    Else
        vHeaderRow = 0
    End If
    
    Exit Sub
ErrorHandle:
    custErr.RaiseError "cRange - DefineRange()"
End Sub


' Properties

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
    LastRow = vRng.Rows(vRng.Rows.Count)
    Exit Property
ErrorHandle:
    custErr.RaiseError "cRange - Get LastRow()"
End Property
