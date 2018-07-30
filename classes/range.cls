Option Explicit

Private vRng As range
Private vHeader As Boolean
Private vHeaderRow As Long



Sub DefineRange(ByRef CellRange As range, Optional HasHeaderRow As Boolean = True, Optional HeaderRow As Long = 1)
    On Error GoTo ErrorHandle
    
    Set vRng = CellRange
    vHeader = HasHeaderRow
    If vHeader Then
        vHeaderRow = HeaderRow
    Else
        vHeaderRow = 0
    End If
    
    Exit Sub
ErrorHandle:
    custErr.RaiseError "cRange - DefineRange()"
End Sub


' Properties

Property Get rng() As range
    On Error GoTo ErrorHandle
    If vRng Is Nothing Then Err.Raise custErr.GenericError, Description:="The range has not been defined"
    Set rng = vRng
    Exit Property
ErrorHandle:
    custErr.RaiseError "cRange - Get rng()"
End Property

