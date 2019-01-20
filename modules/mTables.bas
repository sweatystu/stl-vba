Option Explicit

' Dependencies:
'   - cTable

' Public Procedures
Sub DeleteAllRows()
    ' Description: Delete all rows in the selected table
    ' Dependencies:
    '   - cTable
    ' Inputs: None
    ' Outputs: None
    Dim ws As Worksheet
    Dim cl As Range
    Dim tbl As ListObject
    Dim t As New cTable
    On Error GoTo ErrorHandle
    Set ws = ActiveSheet
    Set cl = ActiveCell
    For Each tbl In ws.ListObjects
        If Not Intersect(cl, tbl.Range) Is Nothing Then GoTo TblFound
    Next tbl
    Err.Raise custErr.GenericError, Description:="Table was not selected"
    Exit Sub
TblFound:
    Set t.lo = tbl
    t.DeleteTableRows
    Exit Sub
ErrorHandle:
    custErr.DisplayError "mTables - DeleteAllRows()"
End Sub
