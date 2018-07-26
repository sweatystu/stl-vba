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
- **Inputs:**
- **Actions:**
- **Outputs:**

### SetOriginalSettings()
- **Prerequisits:** None
- **Inputs:**
- **Actions:**
- **Outputs:**

### PreviousAllSettings()
- **Prerequisits:** None
- **Inputs:**
- **Actions:**
- **Outputs:**

### SetSheetsInNew()
- **Prerequisits:** None
- **Inputs:**
- **Actions:**
- **Outputs:**

### PreviousSheetsInNew()
- **Prerequisits:** None
- **Inputs:**
- **Actions:**
- **Outputs:**

### SetCalculationMode()
- **Prerequisits:** None
- **Inputs:**
- **Actions:**
- **Outputs:**

### PreviousCalculationMode()
- **Prerequisits:** None
- **Inputs:**
- **Actions:**
- **Outputs:**

### SetEvents()
- **Prerequisits:** None
- **Inputs:**
- **Actions:**
- **Outputs:**

### PreviousEvents()
- **Prerequisits:** None
- **Inputs:**
- **Actions:**
- **Outputs:**

### SetAlerts()
- **Prerequisits:** None
- **Inputs:**
- **Actions:**
- **Outputs:**

### PreviousAlerts()
- **Prerequisits:** None
- **Inputs:**
- **Actions:**
- **Outputs:**

### SetScreenUpdate()
- **Prerequisits:** None
- **Inputs:**
- **Actions:**
- **Outputs:**

### PreviousScreenUpdate()
- **Prerequisits:** None
- **Inputs:**
- **Actions:**
- **Outputs:**

## Private Procedures
These procedures are only accessible to other procedures within the class.

### SetDefaultSettings()
- **Prerequisits:** None
- **Inputs:**
- **Actions:**
- **Outputs:**

### Class_Initialize()
- **Prerequisits:** None
- **Inputs:**
- **Actions:**
- **Outputs:**

### Class_Terminate()
- **Prerequisits:** None
- **Inputs:**
- **Actions:**
- **Outputs:**

### RecordOriginalSettings()
- **Prerequisits:** None
- **Inputs:**
- **Actions:**
- **Outputs:**


