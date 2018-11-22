Option Explicit

' Dependencies:

' Constants
Private Const sht_name As String = "#SheetSettings#"
Private Const sht_cl As String = "A1"

' Public Procedures
Sub all_sheets_visible()
' Description: Makes all sheets in the active workbook visible
' Dependencies: None
' Inputs: None
' Outputs: None
    Dim sht As Worksheet
    For Each sht In ActiveWorkbook.Sheets
        sht.Visible = xlSheetVisible
    Next sht
End Sub

Sub toggle_sheets()
' Description: Make all sheets visible and return them to original setting when run again
' Dependencies:
'   - sht_name (const, As String)
' Inputs: None
' Outputs: None
    Dim wb As Workbook
    Dim ws As Worksheet
    Set wb = ActiveWorkbook
    app.Initialise
    On Error Resume Next
    Set ws = wb.Sheets(sht_name)
    If Err.Number = 0 Then
        ' Sheet found - return sheets
        return_sheet_status wb, ws
        MsgBox "Sheets have been returned to original status", vbOKOnly + vbInformation, "Sheets Toggled"
    Else
        ' Sheet not found - show all sheets
        record_sheet_status wb
    End If
End Sub

' Private Procedures
Private Sub record_sheet_status(ByRef wb As Workbook)
' Description: Records the status of all sheets in the workbook and records them in a new sheet
' Dependencies:
'   - sht_name (const, As String)
'   - sht_cl (const, As String)
' Inputs:
'   wb  (As Workbook)   The workbook containing the sheets to record the status of
' Outputs: None
    Dim record As Worksheet
    Dim ws As Worksheet
    Dim cl As Range
    Dim i As Long
    Set record = wb.Sheets.Add
    record.Name = sht_name
    ' record.Visible = xlSheetVeryHidden
    Set cl = record.Range(sht_cl)
    i = 0
    For Each ws In wb.Sheets
        cl.Offset(i, 0).Value = ws.Name
        cl.Offset(i, 1).Value = ws.Visible
        ws.Visible = xlSheetVisible
        i = i + 1
    Next ws
    record.Visible = xlSheetVeryHidden
End Sub

Private Sub return_sheet_status(ByRef wb As Workbook, ByRef record As Worksheet)
' Description: Return sheets to original setting
' Dependencies:
'   - sht_cl  (const, As String)
' Inputs:
'   - wb        (As Workbook)   The workbook containing the sheets to act on
'   - record    (As Worksheet)  The worksheet containing the original sheet names and status
' Outputs: None
    Dim cl As Range
    Dim i As Long
    Set cl = record.Range(sht_cl)
    i = 0
    On Error Resume Next ' Continue if sheet not found
    Do Until cl.Offset(i, 0).Value = ""
        wb.Sheets(cl.Offset(i, 0).Value).Visible = cl.Offset(i, 1).Value
        i = i + 1
    Loop
    record.Delete
End Sub
