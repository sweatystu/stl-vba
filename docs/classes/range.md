# Range

- **Filename:** range.cls
- **Instance Name:** cRange
- **Prerequisits:**
    - datarange.cls - `cDataRange`
    - headerrange.cls - `cHeaderRange`
    - custom-errors.cls - `cCustomErrors` (implemented through *m-global.bas*)

This class is to be used as an extension of the range class. It can be used to interact with ranges in a defined and flexible way, making use of a header row to find columns by name if appropriate.

## General Use
This class must be initialised to be used.

``` VB
Dim rng As New cRange
rng.DefineRange Range("A1:C3"), True, 1
```

This class allows a range to be used with or without a header row. If a header row is used, only the rows *below* the header are considered to be data. This may be useful when importing sheets of data that have metadata in the first few rows.

## Public Procedures

### DefineRange()
- **Prerequisits:**
    - `cDataRange`
    - `cHeaderRange`
- **Inputs:**
    - CellRange As Range - *The whole data range*
    - *Optional* HasHeaderRow As Boolean - *Whether the range has a header row or not, defaults to TRUE*
    - *Optional* HeaderRow As Long - *Data row (not sheet row) that the header is in, defaults to 1*
- **Actions:**
    - **With HeaderRow**
        - Confirm header row is within data range with data below it
        - Set header row as new `cHeaderRange`
        - Set range below header as new `cDataRange`
    - **Without HeaderRow**
        - Set *CellRange* as new `cDataRange`
- **Outputs:** None

## Properties


