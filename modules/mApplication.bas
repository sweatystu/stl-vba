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
    On Error GoTo ErrorHandle
    continue = False
    While Not continue
        continue = pvtChangeSheetsInNew(InputBox("How many sheets should be in a new workbook?", "New Workbook Sheets", Application.SheetsInNewWorkbook))
        If Not continue Then
            If MsgBox("A numeric value must be entered", vbOKCancel + vbInformation, "Insufficient data") = vbCancel Then Exit Sub
        End If
    Wend
    Exit Sub
ErrorHandle:
    custErr.DisplayError "mApplication - ChangeSheetsInNew()"
End Sub

' Private Functions
Private Function pvtChangeSheetsInNew(ByVal vl As String) As Boolean
    ' Description: Daughter Procedure
    '               Changes the number of sheets in a new workbook to the given value
    ' Dependencies:
    '   - cAppProperties
    ' Inputs:
    '   - (As String)   vl  - Output of an InputBox describing the number of sheets to include in a new workbook
    '                       - Value must be able to convert to an integer number
    ' Outputs:
    '   - (As Boolean) Whether the function has completed or not
    On Error GoTo ErrorHandle
    If Len(vl) < 1 Then GoTo FailFunction
    If Not IsNumeric(vl) Then GoTo FailFunction
    app.SetSheetsInNew CInt(vl)
    pvtChangeSheetsInNew = True
    Exit Function
FailFunction:
    pvtChangeSheetsInNew = False
    Exit Function
ErrorHandle:
    custErr.RaiseError "mApplication - pvtChangeSheetsInNew()"
End Function
