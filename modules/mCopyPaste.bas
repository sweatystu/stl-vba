Option Explicit

' Dependencies:
'   - cRange class
'   - cTable class
'   - There must be a UserForm within the project

' Public Procedures
Sub TransposeValues()
    ' Description: Transposes clipboard and pastes as values
    ' Dependencies:
    '   - Must be a range of cells copied to the clipboard
    ' Inputs: None
    ' Outputs: None
    On Error GoTo ErrorHandle
    ActiveCell.PasteSpecial Paste:=xlPasteValues, Operation:=xlPasteSpecialOperationNone, SkipBlanks:=False, Transpose:=True
    Exit Sub
ErrorHandle:
    custErr.DisplayError "mCopyPaste - TransposeValues()"
End Sub

Sub CopyToCSVDefault()
    ' Description: Copies selected cells to clipboard separated by ", "
    ' Dependencies:
    '   - pvt ContentsToCSV()
    ' Inputs: None
    ' Outputs:
    '   - Contents of selected cells copied to clipboard
    Dim Obj As New MSForms.DataObject
    On Error GoTo ErrorHandle
    Application.CutCopyMode = False
    With Obj
        .SetText ContentsToCSV(Selection, ", ")
        .PutInClipboard
    End With
    Exit Sub
ErrorHandle:
    custErr.DisplayError "mCopyPaste - CopyToCSVDefault()"
End Sub

Sub CopyToCSVSelectSep()
    ' Description: Copies selected cells to clipboard separated by provided text
    ' Dependencies:
    '   - pvt ContentsToCSV()
    ' Inputs: None
    ' Outputs:
    '   - Contents of selected cells copied to clipboard
    Dim Obj As New MSForms.DataObject
    On Error GoTo ErrorHandle
    Application.CutCopyMode = False
    With Obj
        .SetText ContentsToCSV(Selection, InputBox("What separator should separate the values?", "Choose Separator", ", "))
        .PutInClipboard
    End With
    Exit Sub
ErrorHandle:
    custErr.DisplayError "mCopyPaste - CopyToCSVSelectSep()"
End Sub

Sub ExportSelection()
    ' Description: Exports the current selection to a new workbook as values
    ' Dependencies:
    '   - cWorkbook class
    ' Inputs: None
    ' Outputs:
    '   - New workbook with values
    Dim wb As New cWorkbook
    Dim r As Range
    On Error GoTo ErrorHandle
    app.Initialise
    Set r = Selection.SpecialCells(xlCellTypeVisible)
    wb.NewWorkbook
    ExportRange r, wb.wb.Sheets(1).Range("A1")
    Exit Sub
ErrorHandle:
    custErr.DisplayError "mCopyPaste - ExportSelection()"
End Sub

Sub ExportTable()
    ' Description: Exports the currently selected table to a new workbook as a range of values
    ' Dependencies:
    '   - cWorkbook class
    '   - cTable class
    ' Inputs: None
    ' Outputs:
    '   - New workbook with values
    Dim wb As New cWorkbook
    Dim ws As Worksheet
    Dim lo As ListObject
    Dim tbl As New cTable
    On Error GoTo ErrorHandle
    app.Initialise
    Set ws = ActiveSheet
    If ws.ListObjects.Count = 0 Then Err.Raise custErr.GenericError, Description:="There are no tables in the current worksheet"
    For Each lo In ws.ListObjects
        If Not Intersect(ActiveCell, lo.Range) Is Nothing Then
            Set tbl.lo = lo
            Exit For
        End If
    Next lo
    If tbl.lo Is Nothing Then Err.Raise custErr.GenericError, Description:="The active cell is not in a table"
    wb.NewWorkbook
    ExportRange tbl.lo.Range, wb.wb.Sheets(1).Range("A1")
    Exit Sub
ErrorHandle:
    custErr.DisplayError "mCopyPaste - ExportTable()"
End Sub

' Private Functions
Private Function ContentsToCSV(ByRef rng As Range, ByVal sep As String) As String
    ' Description: Converts the contents of a range of cells to string separated by the given text
    ' Dependencies: None
    ' Inputs:
    '   - (As Range)    rng     - The range of cells to take the content from
    '   - (As String)   sep     - The separator to place between the contents of each cell
    ' Outputs:
    '   - (As String) The contents of each cell converted to a single string, separated by the given text
    Dim cl As Range
    Dim txt As String
    On Error GoTo ErrorHandle
    For Each cl In rng.SpecialCells(xlCellTypeVisible)
        If Len(cl.Value) > 0 Then
            txt = txt & CStr(cl.Value) & sep
        End If
    Next cl
    If Len(txt) > 0 Then txt = Left(txt, Len(txt) - Len(sep))
    ContentsToCSV = txt
    Exit Function
ErrorHandle:
    custErr.RaiseError "mCopyPaste - ContentsToCSV()"
End Function

' Private Procedures
Private Sub ExportRange(ByRef rng As Range, ByRef target As Range)
    ' Description: Copies the contents of rng to the target cell
    ' Dependencies: None
    ' Inputs:
    '   - (As Range)    rng     - The selection of cells to copy
    '   - (As Range)    target  - The cell that the cells should be copied to
    ' Outputs: None
    On Error GoTo ErrorHandle
    Application.CutCopyMode = False
    rng.Copy
    With target
        .PasteSpecial xlPasteColumnWidths
        .PasteSpecial xlPasteValuesAndNumberFormats
        .PasteSpecial xlPasteFormats
    End With
    Application.CutCopyMode = False
    Exit Sub
ErrorHandle:
    custErr.RaiseError "mCopyPaste - ExportRange()"
End Sub
