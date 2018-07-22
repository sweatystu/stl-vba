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
End Sub



' Initialisation
Private Sub Class_Initialize()
' Record original status of application variables
    Me.RecordOriginalSettings
' Apply predetermined variables to speed up macro running
    Me.SetDefaultSettings
End Sub

' Termination
Private Sub Class_Terminate()
' Remove variables to return to original state
    Me.SetOriginalSettings
End Sub


' Record the original settings
Private Sub RecordOriginalSettings()
    vOrigNoSheets = Application.SheetsInNewWorkbook
    vOrigCalcOptions = Application.Calculation
    vOrigEvents = Application.EnableEvents
    vOrigAlerts = Application.DisplayAlerts
    vOrigScreenUpdate = Application.ScreenUpdating
End Sub

' Apply original settings
Sub SetOriginalSettings()
    Me.SetSheetsInNew vOrigNoSheets
    Me.SetCalculationMode vOrigCalcOptions
    Me.SetEvents vOrigEvents
    Me.SetAlerts vOrigAlerts
    Me.SetScreenUpdate vOrigScreenUpdate
End Sub

' Roll back all settings to previous
Sub PreviousAllSettings()
    Me.PreviousSheetsInNew
    Me.PreviousCalculationMode
    Me.PreviousAlerts
    Me.PreviousEvents
    Me.PreviousScreenUpdate
End Sub


' Sheets in New Workbook
Sub SetSheetsInNew(ByVal i As Long)
    ' Record Previous sheets in workbook
    vPrevNoSheets = Application.SheetsInNewWorkbook
    ' Apply new number of sheets in a workbook
    Application.SheetsInNewWorkbook = i
End Sub

' Previous Sheets in New Workbook
Sub PreviousSheetsInNew()
    Me.SetSheetsInNew vPrevNoSheets
End Sub


' Calculation Mode
Sub SetCalculationMode(ByVal vl As XlCalculation)
    ' Record previous calculation mode
    vPrevCalcOptions = Application.Calculation
    ' Apply new calculation mode
    Application.Calculation = vl
End Sub

' Previous Calculation Mode
Sub PreviousCalculationMode()
    Me.SetCalculationMode vPrevCalcOptions
End Sub


' Events
Sub SetEvents(ByVal TF As Boolean)
    ' Record previous event status
    vPrevEvents = Application.EnableEvents
    ' Apply new event status
    Application.EnableEvents = TF
End Sub

' Previous Event Mode
Sub PreviousEvents()
    Me.SetEvents vPrevEvents
End Sub


' Alerts
Sub SetAlerts(ByVal TF As Boolean)
    ' Record previous alert status
    vPrevAlerts = Application.DisplayAlerts
    ' Apply new alert status
    Application.DisplayAlerts = TF
End Sub

' Previous Alert Mode
Sub PreviousAlerts()
    Me.SetAlerts vPrevAlerts
End Sub


' Screen Updating
Sub SetScreenUpdate(ByVal TF As Boolean)
    ' Record previous screen updating status
    vPrevScreenUpdate = Application.ScreenUpdating
    ' Apply new screen updating status
    Application.ScreenUpdating = TF
End Sub

' Previous Screen Updating Mode
Sub PreviousScreenUpdate()
    Me.SetScreenUpdate vPrevScreenUpdate
End Sub
