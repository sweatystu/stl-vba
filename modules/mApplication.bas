Option Explicit

' Dependencies:
'   - cAppProperties

' Public Procedures
Sub ToggleEvents()
    ' Description: Toggle between events on and off
    ' Dependencies:
    '   - cAppProperties
    ' Inputs: None
    ' Outputs: None
    Dim txt As String
    On Error GoTo ErrorHandle
    app.SetEvents Not Application.EnableEvents
    If Application.EnableEvents Then
        txt = "Events have been enabled"
    Else
        txt = "Events have been disabled"
    End If
    MsgBox txt, VbMsgBoxStyle.vbOKOnly + VbMsgBoxStyle.vbInformation, "Events Toggled"
    Exit Sub
ErrorHandle:
    custErr.DisplayError "mApplication - ToggleEvents()"
End Sub

Sub ToggleCalculationMode()
    ' Description: Toggles the calculation mode between manual and automatic
    ' Dependencies:
    '   - cAppProperties
    ' Inputs: None
    ' Outputs: None
    Dim txt As String
    On Error GoTo ErrorHandle
    Select Case Application.Calculation
        Case XlCalculation.xlCalculationAutomatic
            app.SetCalculationMode XlCalculation.xlCalculationManual
            txt = "Calculation mode has been set to manual"
        Case Else
            app.SetCalculationMode XlCalculation.xlCalculationAutomatic
            txt = "Calculation mode has been set to automatic"
    End Select
    MsgBox txt, VbMsgBoxStyle.vbOKOnly + VbMsgBoxStyle.vbInformation, "Calculation Mode Toggled"
    Exit Sub
ErrorHandle:
    custErr.DisplayError "mApplication - ToggleCalculationMode()"
End Sub

Sub ChangeSheetsInNew()
    ' Description: Mother Procedure
    '               Changes the number of sheets in a new workbook to the given value
    ' Dependencies:
    '   - Me.pvtChangeSheetsInNew()
    ' Inputs: None
    ' Outputs: None
    Dim continue As Boolean
    Dim vl As String
    On Error GoTo ErrorHandle
    continue = False
    vl = ""
    While Not continue
        vl = InputBox("How many sheets should be in a new workbook?", "New Workbook Sheets", Application.SheetsInNewWorkbook)
        If Len(vl) < 1 Then Exit Sub
        If Not IsNumeric(vl) Then
            MsgBox "The value entered must be numeric", vbOKOnly + vbInformation, "Value Error"
        Else
            app.SetSheetsInNew CLng(vl)
            continue = True
        End If
    Wend
    Exit Sub
ErrorHandle:
    custErr.DisplayError "mApplication - ChangeSheetsInNew()"
End Sub
