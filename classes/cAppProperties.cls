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



' Initialisation
Sub Initialise()
' Empty procedure
' Call this procedure to force the class to initialise
    Class_Initialize
End Sub

Private Sub Class_Initialize()
    On Error GoTo ErrorHandle
    
' Record original status of application variables
    RecordOriginalSettings
' Apply predetermined variables to speed up macro running
    SetDefaultSettings
    Exit Sub
    
ErrorHandle:
    custErr.RaiseError "cAppProperties - Class_Initialize()"
End Sub

' Termination
Private Sub Class_Terminate()
    On Error GoTo ErrorHandle
    
' Remove variables to return to original state
    SetOriginalSettings
    Exit Sub
    
ErrorHandle:
    custErr.RaiseError "cAppProperties - Class_Terminate()"
End Sub


' Record the original settings
Private Sub RecordOriginalSettings()
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

' Apply original settings
Sub SetOriginalSettings()
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

' Roll back all settings to previous
Sub PreviousAllSettings()
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


' Sheets in New Workbook
Sub SetSheetsInNew(ByVal i As Long)
    On Error GoTo ErrorHandle
    
    ' Record Previous sheets in workbook
    vPrevNoSheets = Application.SheetsInNewWorkbook
    ' Apply new number of sheets in a workbook
    Application.SheetsInNewWorkbook = i
    
    Exit Sub
    
ErrorHandle:
    custErr.RaiseError "cAppProperties - SetSheetsInNew()"
End Sub

' Previous Sheets in New Workbook
Sub PreviousSheetsInNew()
    On Error GoTo ErrorHandle
    
    Me.SetSheetsInNew vPrevNoSheets
    Exit Sub
    
ErrorHandle:
    custErr.RaiseError "cAppProperties - PreviousSheetsInNew()"
End Sub


' Calculation Mode
Sub SetCalculationMode(ByVal vl As XlCalculation)
    On Error GoTo ErrorHandle
    
    ' Record previous calculation mode
    vPrevCalcOptions = Application.Calculation
    ' Apply new calculation mode
    Application.Calculation = vl
    
    Exit Sub
    
ErrorHandle:
    custErr.RaiseError "cAppProperties - SetCalculationMode()"
End Sub

' Previous Calculation Mode
Sub PreviousCalculationMode()
    On Error GoTo ErrorHandle
    
    Me.SetCalculationMode vPrevCalcOptions
    Exit Sub
    
ErrorHandle:
    custErr.RaiseError "cAppProperties - PreviousCalculationMode()"
End Sub


' Events
Sub SetEvents(ByVal TF As Boolean)
    On Error GoTo ErrorHandle
    
    ' Record previous event status
    vPrevEvents = Application.EnableEvents
    ' Apply new event status
    Application.EnableEvents = TF
    
    Exit Sub
    
ErrorHandle:
    custErr.RaiseError "cAppProperties - SetEvents()"
End Sub

' Previous Event Mode
Sub PreviousEvents()
    On Error GoTo ErrorHandle
    
    Me.SetEvents vPrevEvents
    Exit Sub
    
ErrorHandle:
    custErr.RaiseError "cAppProperties - PreviousEvents()"
End Sub


' Alerts
Sub SetAlerts(ByVal TF As Boolean)
    On Error GoTo ErrorHandle
    
    ' Record previous alert status
    vPrevAlerts = Application.DisplayAlerts
    ' Apply new alert status
    Application.DisplayAlerts = TF
    
    Exit Sub
    
ErrorHandle:
    custErr.RaiseError "cAppProperties - SetAlerts()"
End Sub

' Previous Alert Mode
Sub PreviousAlerts()
    On Error GoTo ErrorHandle
    
    Me.SetAlerts vPrevAlerts
    Exit Sub
    
ErrorHandle:
    custErr.RaiseError "cAppProperties - PreviousAlerts()"
End Sub


' Screen Updating
Sub SetScreenUpdate(ByVal TF As Boolean)
    On Error GoTo ErrorHandle
    
    ' Record previous screen updating status
    vPrevScreenUpdate = Application.ScreenUpdating
    ' Apply new screen updating status
    Application.ScreenUpdating = TF
    
    Exit Sub
    
ErrorHandle:
    custErr.RaiseError "cAppProperties - SetScreenUpdate()"
End Sub

' Previous Screen Updating Mode
Sub PreviousScreenUpdate()
    On Error GoTo ErrorHandle
    
    Me.SetScreenUpdate vPrevScreenUpdate
    Exit Sub
    
ErrorHandle:
    custErr.RaiseError "cAppProperties - PreviousScreenUpdate()"
End Sub

