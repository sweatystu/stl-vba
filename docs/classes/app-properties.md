# App Properties

- **Filename:** app-properties.cls
- **Instance Name:** cAppProperties
- **Prerequisits:** None

This class is initiated as a global variable (`app`) and can be used in any procedure.

## General Use
If initialised, using `app.Initialise`, at the beginning of an automated process, standard defined settings will be applied. These settings will increase the speed at which an automated process is carried out (*e.g.* screen updating turned off) and standardise settings making it easier to write new macros (*e.g.* number of sheets in a new  workbook set to 1).

``` VB
Sub Example()
    On Error GoTo ErrorHandle
    app.Initialise

    '''' Your Code Here ''''

End Sub
```

## Public Procedures
These procedures can be called by procedures in other modules.

### Initialise()
- **Prerequisits:** None
- **Inputs:** None
- **Actions:**
    - `Class_Initialise()` - Private Procedure to initialise class
- **Outputs:** None

### SetOriginalSettings()
- **Prerequisits:** None
- **Inputs:** None
- **Actions:**
    - Return app settings to original state
- **Outputs:** None

### PreviousAllSettings()
- **Prerequisits:** None
- **Inputs:** None
- **Actions:**
    - Return app settings to the state before their last change through the class
- **Outputs:** None

### SetSheetsInNew()
- **Prerequisits:** None
- **Inputs:**
    - i As Long - *number of sheets in a new workbook*
- **Actions:**
    - Set the number of sheets in a new workbook to *i*
- **Outputs:** None

### PreviousSheetsInNew()
- **Prerequisits:** None
- **Inputs:** None
- **Actions:**
    - Set the number of sheets in a new workbook to the number before the last change through the class
- **Outputs:** None

### SetCalculationMode()
- **Prerequisits:** None
- **Inputs:**
    - vl As xlCalculation - *Calculation mode (xlCalculationManual / xlCalculationAutomatic)*
- **Actions:**
    - Set the calculation mode to *vl*
- **Outputs:** None

### PreviousCalculationMode()
- **Prerequisits:** None
- **Inputs:** None
- **Actions:**
    Set the calculation mode to the value before the last change through the class
- **Outputs:** None

### SetEvents()
- **Prerequisits:** None
- **Inputs:**
    - TF As Boolean - *True or False*
- **Actions:**
    - Set the *Events Enabled* setting to *TF*
- **Outputs:** None

### PreviousEvents()
- **Prerequisits:** None
- **Inputs:** None
- **Actions:**
    - Set the *Events Enabled* setting to the value before the last change through the class
- **Outputs:** None

### SetAlerts()
- **Prerequisits:** None
- **Inputs:**
    - TF As Boolean - *True or False*
- **Actions:**
    - Set the *Display Alerts* setting to *TF*
- **Outputs:** None

### PreviousAlerts()
- **Prerequisits:** None
- **Inputs:** None
- **Actions:**
    - Set the *Display Alerts* setting to the value before the last change through the class
- **Outputs:** None

### SetScreenUpdate()
- **Prerequisits:** None
- **Inputs:**
    - TF As Boolean - *True or False*
- **Actions:**
    - Set the *Screen Updating* setting to *TF*
- **Outputs:** None

### PreviousScreenUpdate()
- **Prerequisits:** None
- **Inputs:** None
- **Actions:**
    - Set the *Screen Updating* setting to the value before the last change through the class
- **Outputs:** None

## Private Procedures
These procedures are only accessible to other procedures within the class.

### SetDefaultSettings()
- **Prerequisits:** None
- **Inputs:** None
- **Actions:**
    - `SetCalculationMode xlCalculationManual` - turn calculations to manual
    - `SetAlerts False` - turn off alerts
    - `SetEvents False` - turn off events
    - `SetSheetsInNew 1` - set the number of sheets in a new workbook to 1
    - `SetScreenUpdate False` - turn off screen updating
- **Outputs:** None

### Class_Initialize()
- **Prerequisits:** None
- **Inputs:** None
- **Actions:**
    - `RecordOriginalSettings` - record the original settings
    - `SetDefaultSettings` - set the default settings to speed up macros
- **Outputs:** None

### Class_Terminate()
- **Prerequisits:** None
- **Inputs:** None
- **Actions:**
    - `SetOriginalSettings` - return settings to their original state
- **Outputs:** None

### RecordOriginalSettings()
- **Prerequisits:** None
- **Inputs:** None
- **Actions:**
    - Record the original settings
        - Number of sheets in a new workbook
        - Calculation mode
        - Events
        - Alerts
        - Screen Updating
- **Outputs:** None


