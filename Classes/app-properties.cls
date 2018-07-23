Option Explicit

' Constants
Private vOrigNoSheets As Long
Private vOrigCalcOptions As XlCalculation
Private vOrigEvents As Boolean
Private vOrigAlerts As Boolean
Private vOrigScreenUpdate As Boolean

Private vPrevNoSheets As Long
Private vPrevCalcOptions As XlCalculation
Private vPrevEvents As Boolean
Private vPrevAlerts As Boolean
Private vPrevScreenUpdate As Boolean



' Default settings to be applied on initialise
' Change this procedure to alter default behaviour
Private Sub SetDefaultSettings()
    Me.SetCalculationMode XlCalculation.xlCalculationManual
    Me.SetAlerts False
    Me.SetEvents True
    Me.SetSheetsInNew 1
    Me.ScreenUpdating False
    Exit Sub
End Sub



' Initialisation
Private Sub Class_Initialize()
' Record original status of application variables
    If Not RecordOriginalSettings Then Err.Raise custErr.ApplicationError, "cls-AppProperties - Class_Initialize()", "The original settings could not be recorded"
' Apply predetermined variables to speed up macro running
    If Not SetDefaultSettings Then Err.Raise custErr.ApplicationError, "cls-AppProperties - Class_Initialize()", "The default settings could not be set"
End Sub

' Termination
Private Sub Class_Terminate()
' Remove variables to return to original state
    If Not SetOriginalSettings Then Err.Raise custError.ApplicationError, "cls-AppProperties - Class_Terminate()", "The original settings could not be returned"
End Sub


' Record the original settings
Private Function RecordOriginalSettings() As Boolean
    vOrigNoSheets = Application.SheetsInNewWorkbook
    vOrigCalcOptions = Application.Calculation
    vOrigEvents = Application.EnableEvents
    vOrigAlerts = Application.DisplayAlerts
    vOrigScreenUpdate = Application.ScreenUpdating

    RecordOriginalSettings = True
End Function

' Apply original settings
Function SetOriginalSettings() As Boolean
    Me.SetSheetsInNew vOrigNoSheets
    Me.SetCalculationMode vOrigCalcOptions
    Me.SetEvents vOrigEvents
    Me.SetAlerts vOrigAlerts
    Me.SetScreenUpdate vOrigScreenUpdate

    SetOriginalSettings = True
End Function

' Roll back all settings to previous
Function PreviousAllSettings() As Boolean
    Me.PreviousSheetsInNew
    Me.PreviousCalculationMode
    Me.PreviousAlerts
    Me.PreviousEvents
    Me.PreviousScreenUpdate

    PreviousAllSettings = True
End Function


' Sheets in New Workbook
Function SetSheetsInNew(ByVal i As Long) As Boolean
    ' Record Previous sheets in workbook
    vPrevNoSheets = Application.SheetsInNewWorkbook
    ' Apply new number of sheets in a workbook
    Application.SheetsInNewWorkbook = i

    SetSheetsInNew = True
End Function

' Previous Sheets in New Workbook
Function PreviousSheetsInNew() As Boolean
    Me.SetSheetsInNew vPrevNoSheets
    PreviousSheetsInNew = True
End Function


' Calculation Mode
Function SetCalculationMode(ByVal vl As XlCalculation) As Boolean
    ' Record previous calculation mode
    vPrevCalcOptions = Application.Calculation
    ' Apply new calculation mode
    Application.Calculation = vl

    SetCalculationMode = True
End Function

' Previous Calculation Mode
Function PreviousCalculationMode() As Boolean
    Me.SetCalculationMode vPrevCalcOptions
    PreviousCalculationMode = True
End Function


' Events
Function SetEvents(ByVal TF As Boolean) As Boolean
    ' Record previous event status
    vPrevEvents = Application.EnableEvents
    ' Apply new event status
    Application.EnableEvents = TF

    SetEvents = True
End Function

' Previous Event Mode
Function PreviousEvents() As Boolean
    Me.SetEvents vPrevEvents
    PreviousEvents = True
End Function


' Alerts
Function SetAlerts(ByVal TF As Boolean) As Boolean
    ' Record previous alert status
    vPrevAlerts = Application.DisplayAlerts
    ' Apply new alert status
    Application.DisplayAlerts = TF

    SetAlerts = True
End Function

' Previous Alert Mode
Function PreviousAlerts() As Boolean
    Me.SetAlerts vPrevAlerts
    PreviousAlerts = True
End Function


' Screen Updating
Function SetScreenUpdate(ByVal TF As Boolean) As Boolean
    ' Record previous screen updating status
    vPrevScreenUpdate = Application.ScreenUpdating
    ' Apply new screen updating status
    Application.ScreenUpdating = TF

    SetScreenUpdate = True
End Function

' Previous Screen Updating Mode
Function PreviousScreenUpdate() As Boolean
    Me.SetScreenUpdate vPrevScreenUpdate
    PreviousScreenUpdate = True
End Function
