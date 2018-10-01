Option Explicit

' Dependencies:
'   - cCustomErrors class

' Private Variables
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

' Private Procedures
Private Sub Class_Terminate()
    ' Description: Script carried out when class is terminated
    '               Returns settings to original states
    ' Dependencies:
    '   - Me.SetOriginalSettings()
    ' Inputs: None
    ' Outputs: None
    On Error GoTo ErrorHandle
    Me.SetOriginalSettings
    Exit Sub
ErrorHandle:
    custErr.RaiseError "cAppProperties - Class_Terminate()"
End Sub

Private Sub SetDefaultSettings()
    ' Description: Applies default settings to speed up a macro
    ' Dependencies: None
    ' Inputs: None
    ' Outputs: None
    On Error GoTo ErrorHandle
    Me.SetCalculationMode XlCalculation.xlCalculationManual
    Me.SetAlerts False
    Me.SetEvents False
    Me.SetSheetsInNew 1
    Me.SetScreenUpdate False
    Exit Sub
ErrorHandle:
    custErr.RaiseError "cAppProperties - SetDefaultSettings()"
End Sub

Private Sub RecordOriginalSettings()
    ' Description: Records original application settings to be reapplied later
    ' Dependencies: None
    ' Inputs: None
    ' Outputs: None
    On Error GoTo ErrorHandle
    vOrigNoSheets = Application.SheetsInNewWorkbook
    vOrigCalcOptions = Application.Calculation
    vOrigEvents = Application.EnableEvents
    vOrigAlerts = Application.DisplayAlerts
    vOrigScreenUpdate = Application.ScreenUpdating
    Exit Sub
ErrorHandle:
    custErr.RaiseError "cAppProperties - RecordOriginalSettings()"
End Sub

' Public Procedures
Sub Initialise()
    ' Description: Script carried out to initialise a class
    '               Records original settings and applies defaults to make macros faster
    ' Dependencies:
    '   - RecordOriginalSettings()  - Private
    '   - SetDefaultSettings()      - Private
    ' Inputs: None
    ' Outputs: None
    On Error GoTo ErrorHandle
    RecordOriginalSettings
    SetDefaultSettings
    Exit Sub
ErrorHandle:
    custErr.RaiseError "cAppProperties - Initialize()"
End Sub

Sub SetOriginalSettings()
    ' Description: Returns all settings to their originally recorded state
    ' Dependencies:
    '   - Class must have been initialised via Initialise()
    '   - Me.SetSheetsInNew()
    '   - Me.SetCalculationMode()
    '   - Me.SetEvents()
    '   - Me.SetAlerts()
    '   - Me.SetScreenUpdate()
    ' Inputs: None
    ' Outputs: None
    On Error GoTo ErrorHandle
    Me.SetSheetsInNew vOrigNoSheets
    Me.SetCalculationMode vOrigCalcOptions
    Me.SetEvents vOrigEvents
    Me.SetAlerts vOrigAlerts
    Me.SetScreenUpdate vOrigScreenUpdate
    Exit Sub
ErrorHandle:
    custErr.RaiseError "cAppProperties - SetOriginalSettings()"
End Sub

Sub PreviousAllSettings()
    ' Description: Returns all settings to their last recorded state
    ' Dependencies:
    '   - Settings must have been previously been changed by this class
    '   - Me.PreviousSheetsInNew()
    '   - Me.PreviousCalculationMode()
    '   - Me.PreviousAlerts()
    '   - Me.PreviousEvents()
    '   - Me.PreviousScreenUpdate()
    ' Inputs: None
    ' Outputs: None
    On Error GoTo ErrorHandle
    Me.PreviousSheetsInNew
    Me.PreviousCalculationMode
    Me.PreviousAlerts
    Me.PreviousEvents
    Me.PreviousScreenUpdate
    Exit Sub
ErrorHandle:
    custErr.RaiseError "cAppProperties - PreviousAllSettings()"
End Sub

Sub SetSheetsInNew(ByVal i As Long)
    ' Description: Sets the number of sheets in a new workbook
    ' Dependencies: None
    ' Inputs:
    '   - (As Long) i   - The number of sheets in a new workbook to be applied
    ' Outputs: None
    On Error GoTo ErrorHandle
    vPrevNoSheets = Application.SheetsInNewWorkbook
    Application.SheetsInNewWorkbook = i
    Exit Sub
ErrorHandle:
    custErr.RaiseError "cAppProperties - SetSheetsInNew()"
End Sub

Sub PreviousSheetsInNew()
    ' Description: Returns the number of sheets in a new workbook to the previously recorded value
    ' Dependencies:
    '   - Number of sheets in new must have been previously set by this class
    '   - Me.SetSheetsInNew()
    ' Inputs: None
    ' Outputs: None
    On Error GoTo ErrorHandle
    Me.SetSheetsInNew vPrevNoSheets
    Exit Sub
ErrorHandle:
    custErr.RaiseError "cAppProperties - PreviousSheetsInNew()"
End Sub

Sub SetCalculationMode(ByVal vl As XlCalculation)
    ' Description: Sets the calculation mode
    ' Dependencies: None
    ' Inputs:
    '   - (As XlCalculation)    vl  - The calculation mode to be applied
    ' Outputs: None
    On Error GoTo ErrorHandle
    vPrevCalcOptions = Application.Calculation
    Application.Calculation = vl
    Exit Sub
ErrorHandle:
    custErr.RaiseError "cAppProperties - SetCalculationMode()"
End Sub

Sub PreviousCalculationMode()
    ' Description: Returns the calculation mode to the previously recorded value
    ' Dependencies:
    '   - Calculation mode must have been previously set by this class
    '   - Me.SetCalculationMode()
    ' Inputs: None
    ' Outputs: None
    On Error GoTo ErrorHandle
    Me.SetCalculationMode vPrevCalcOptions
    Exit Sub
ErrorHandle:
    custErr.RaiseError "cAppProperties - PreviousCalculationMode()"
End Sub

Sub SetEvents(ByVal TF As Boolean)
    ' Description: Sets whether events should be triggered or not
    ' Dependencies: None
    ' Inputs:
    '   - (As Boolean)  TF  - Whether events should be triggered or not
    ' Outputs: None
    On Error GoTo ErrorHandle
    vPrevEvents = Application.EnableEvents
    Application.EnableEvents = TF
    Exit Sub
ErrorHandle:
    custErr.RaiseError "cAppProperties - SetEvents()"
End Sub

Sub PreviousEvents()
    ' Description: Returns the events to the previously recorded value
    ' Dependencies:
    '   - Events mode must have been previously set by this class
    '   - Me.SetEvents()
    ' Inputs: None
    ' Outputs: None
    On Error GoTo ErrorHandle
    Me.SetEvents vPrevEvents
    Exit Sub
ErrorHandle:
    custErr.RaiseError "cAppProperties - PreviousEvents()"
End Sub

Sub SetAlerts(ByVal TF As Boolean)
    ' Description: Sets whether alerts should be triggered or not
    ' Dependencies: None
    ' Inputs:
    '   - (As Boolean)  TF  - Whether alerts should be triggered or not
    ' Outputs: None
    On Error GoTo ErrorHandle
    vPrevAlerts = Application.DisplayAlerts
    Application.DisplayAlerts = TF
    Exit Sub
ErrorHandle:
    custErr.RaiseError "cAppProperties - SetAlerts()"
End Sub

Sub PreviousAlerts()
    ' Description: Returns the alerts to the previously recorded value
    ' Dependencies:
    '   - Alerts mode must have been previously set by this class
    '   - Me.SetAlerts()
    ' Inputs: None
    ' Outputs: None
    On Error GoTo ErrorHandle
    Me.SetAlerts vPrevAlerts
    Exit Sub
ErrorHandle:
    custErr.RaiseError "cAppProperties - PreviousAlerts()"
End Sub

Sub SetScreenUpdate(ByVal TF As Boolean)
    ' Description: Sets whether screen-updating should occur or not
    ' Dependencies: None
    ' Inputs:
    '   - (As Boolean)  TF  - Whether screen-updating should occur or not
    ' Outputs: None
    On Error GoTo ErrorHandle
    vPrevScreenUpdate = Application.ScreenUpdating
    Application.ScreenUpdating = TF
    Exit Sub
ErrorHandle:
    custErr.RaiseError "cAppProperties - SetScreenUpdate()"
End Sub

Sub PreviousScreenUpdate()
    ' Description: Returns screen-updating to the previously recorded value
    ' Dependencies:
    '   - Screen-updating mode must have been previously set by this class
    '   - Me.SetScreenUpdate()
    ' Inputs: None
    ' Outputs: None
    On Error GoTo ErrorHandle
    Me.SetScreenUpdate vPrevScreenUpdate
    Exit Sub
ErrorHandle:
    custErr.RaiseError "cAppProperties - PreviousScreenUpdate()"
End Sub

