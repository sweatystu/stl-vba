# Custom Errors

- **Filename:** custom-errors.cls
- **Instance Name:** cCustomErrors
- **Prerequisits**: None

This class is initiated as a global variable (`custErr`) and can be used in any procedure.

## General Use

There are two categories of procedure:
- Top-Level
- Sub-Level

*Top Level* procedures act as an entry to the automated process and call *Sub Level* procedures, but do little else. All the work should be carried out by *Sub Level* procedures.

### Top Level

*Top Level* procedures should contain the following code:

``` VB
Sub ExampleTopLevel()
    On Error GoTo ErrorHandle

    '''' Your Code Here ''''

    Exit Sub
ErrorHandle:
    custErr.DisplayError "ModuleName - ExampleTopLevel()"
End Sub
```

When an error is triggered the code immediately goes to the `ErrorHandle` line and envokes the `DisplayError` procedure to display details of the error. It is important to have the `Exit Sub` line before the `ErrorHandle` line.

### Sub Level

The following example of a *Sub Level* procedure includes a line to throw a custom error.

``` VB
Sub ExampleSubLevel()
    On Error GoTo ErrorHandle

    ''''' Your Code Here ''''

    ' Custom Error
    Err.Raise custErr.GenericError, Description:="Test Error"

    '''' Your Code Here ''''

    Exit Sub
ErrorHandle:
    custErr.RaiseError "ModuleName - ExampleSubLevel()"
End Sub
```

When an error is triggered (automatically or manually) the code immediately goes to the `ErrorHandle` line and envokes the `RaiseError` procedure. This raises an error, recording the source of the error for debugging purposes, and propogates the error to the *Top Level* procedure, as long as all *Sub Level* procedures are written following this error handling pattern.

## Properties

Property Name | Description
---- | ----
custErr.GenericError | A generic error that doesn't fit into any other defined category

## Procedures

### RaiseError()
- **Prerequisits:** None
- **Inputs:**
    - procedure As String - *Name of the procedure the procedure was called from*
- **Actions:**
    - Add `procedure` to the `Err.Source` variable
    - Adds the line number to the `Err.Source` if available
- **Outputs:**
    - An `error` with updated `Err.Source`

### DisplayError()
- **Prerequisits:** None
- **Inputs:**
    - procedure As String - *Name of the (Top Level) procedure the procedure was called from
- **Actions:**
    - Produce string containing details of the raised error, including category (Number), source, and description
- **Outputs:**
    - MessageBox containing details of the raised error
