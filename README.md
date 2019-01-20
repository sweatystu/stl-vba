# stl-vba
Collection of VBA scripts

## Introduction
This collection of scripts is intended as a private collection of VBA scripts that have proven useful. They involve a number of classes with useful generic functions, and some procedures that make use of these classes.

## Installation
It is best to install all classes and modules.

All code is written assuming that the `cAppProperties` class and the `cCustomErrors` class have been installed, and that the `mGlobal` module has been installed to automatically initiate instances of the two classes.

## Writing Procedures
Procedures should be written in a modular way with 1 *mother* procedure (the one activated by the user) and any number of *daughter* procedures (called by another procedure).

The *daughter* procedures should carry out all processing and in the event of an error, throw an error which is then propagated up to the *mother* procedure.

``` VB
Sub daughterProcedure()
    ' Description: Sample Procedure
    ' Dependencies: List the 
    '   - classes
    '   - modules or
    '   - daughter procedures used
    ' Inputs: None
    ' Outputs: None
    On Error GoTo ErrorHandle

    ''''' Your code here '''''

    ' Throw an error if required
    If True Then Err.Raise custError.GenericError, Description:="This error has been raised as an example"

    ''''' More code here '''''

    Exit Sub ' End Procedure
ErrorHandle:
    custError.RaiseError "Module Name - ProcedureName()" ' Throw error and pass details to mother procedure
End Sub
```

The *mother* procedure should contain only calls to other procedures and should do no processing of its own. This procedure should define error handling capabilities and optimise settings if required.

``` VB
Sub motherProcedure()
    ' Description: Sample Procedure
    ' Dependencies: List the 
    '   - classes
    '   - modules or
    '   - daughter procedures used
    ' Inputs: None
    ' Outputs: None
    On Error GoTo ErrorHandle
    app.Initialise ' Optimise settings for a fast macro
    
    ''''' Call various procedures and functions '''''

    Exit Sub ' End procedure
ErrorHandle:
    custError.DisplayError "Module Name - ProcedureName()" ' Show recorded error and location in script
End Sub
```

If the macro is relatively long and would benefit from displaying progress to the user, the progress form can be used.

``` VB
Sub motherProcedure()
    ' Description: Sample Procedure with Progress Form
    ' Dependencies:
    '   - cProgress class
    '   - Other modules and classes required
    ' Inputs: None
    ' Outputs: None
    On Error GoTo ErrorHandle
    app.Initialise ' Optimise settings for a fast macro
    progress.LoadForm "Test Macro", "Example macro to show how code could be written"

    progress.AddTask "Task 1"
    ''''' Your code here '''''
    progress.CompleteTask

    progress.AddTask "Task 2"
    ''''' Your code here '''''
    If complete Then
        progress.CompleteTask
    Else
        progress.FailTask
    End If

    progress.CompleteMacro
    Exit Sub
ErrorHandle:
    progress.DisplayError "Example - motherProcedure()"
End Sub
```

## Classes

### cAppProperties

Initialised via `mGlobal` module. Run `Initialise` procedure at beginning of a script to force optimised settings to be applied.

``` VB
app.Initialise
```

### cColours

*Usually* initialised within a class, but can be initialised as required.

``` VB
Dim col As New cColours
```

### cCustomErrors

Initialised via `mGlobal` module. In *daughter* procedures, use `custError.RaiseError` procedure to propagate an error to the *mother* procedure. In the *mother* procedure, use `custError.DisplayError` procedure to display the error and its source. *See examples above*.

### cDataRange

Initialised in `cRange` class.

### cGraph
Needs to be initialised and then the *ChartObject* passed via the `DefineGraph` procedure.

``` VB
Dim g As New cGraph
g.DefineGraph ActiveChart.Parent
' or...
g.DefineGraph ActiveSheet.ChartObjects(1)
```

### cGrey

Initialised in `cColours` class.

### cHeaderRange

Initialised in `cRange` class.

### cIndividualColour

Initialised in `cColours` class.

### cPivot
Needs to be initialised and then the pivot table (as a PivotTable object) passed via the `Pivot` property.

``` VB
Dim pvt As New cPivot
Set pvt.Pivot = ActiveSheets.PivotTables(1)
```

### cProgress

Class is initialised via the `mGlobal` module as the variable `progress`.

Form needs to be loaded by running the `LoadForm` procedure and passing the required arguments:
- **Title** - Text to be used for the title of the macro
- **Description** - Text to be used for the description of the macro

``` VB
progress.LoadForm "MacroTitle", "MacroDescription"

progress.AddTask "Task Description"
''''' Your code here '''''
progress.CompleteTask

progress.CompleteMacro
Exit Sub
ErrorHandle:
progress.DisplayError "ModuleName - ProcedureName()"
```

### cRange

Class needs to be initialised by running the `DefineRange` procedure and passing the required arguments:
- **Range** - Range of cells to be defined as the range
- **Header** - *Optional* - True/False as to whether the range has a header row or not. Default is True.
- **HeaderRow** - *Optional* - Data row (not the sheet row) that the header is in. Default is 1.
    - If the selected range begins in sheet row 4, but the headers are shown in sheet row 6, the *HeaderRow* would be 3 as row 6 is the 3rd data row.

``` VB
Dim rng As New cRange
rng.DefineRange Range("A1:C3"), True, 1
```

### cTable

Class needs to be initialised and then the table (listobject) passed via the `lo` property.

``` VB
Dim tbl As New cTable
Set tbl.lo = ActiveSheet.ListObjects(1)
```

### cWorkbook

Class needs to be initialised and then either a new workbook created or an existing workbook opened.

``` VB
Dim wb1 As New cWorkbook
Dim wb2 As New cWorkbook
wb1.NewWorkbook
wb2.OpenWorkbook "C:\wrkbk.xlsx"
```
