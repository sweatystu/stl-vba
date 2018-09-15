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

## Properties (Get)

### rng()
- **Prerequisits:**
    - The range must be defined by `DefineRange()`
- **Inputs:** None
- **Actions:** None
- **Outputs:**
    - The whole range defined in `DefineRange()` - `Range`

### Sheet()
- **Prerequisits:**
    - The range must be defined by `DefineRange()`
- **Inputs:** None
- **Actions:** None
- **Outputs:**
    - The sheet the defined range is in - `Worksheet`

### wb()
- **Prerequisits:**
    - The range must be defined by `DefineRange()`
- **Inputs:** None
- **Actions:** None
- **Outputs:**
    - The workbook the defined range is in - `Workbook`

### HeaderRange()
- **Prerequisits:**
    - The range must be defined by `DefineRange()` with a header row
    - `cHeaderRange` class
- **Inputs:** None
- **Actions:** None
- **Outputs:**
    - The header range - `cHeaderRange`

### DataRange()
- **Prerequisits:**
    - The range must be defined by `DefineRange()`
    - `cDataRange` class
- **Inputs:** None
- **Actions:** None
- **Outputs:**
    - **No Header:** The range defined by `DefineRange()` - `cDataRange`
    - **With Header:** The range defined by `DefineRange()` underneath the header row - `cDataRange`

### NoColumns()
- **Prerequisits:**
    - The range must be defined by `DefineRange()`
- **Inputs:** None
- **Actions:** None
- **Outputs:**
    - The number of columns in the range - `Long`

### FirstColNum()
- **Prerequisits:**
    - The range must be defined by `DefineRange()`
- **Inputs:** None
- **Actions:** None
- **Outputs:**
    - The sheet number of the first column in the data range - `Long`

### LastColNum()
- **Prerequisits:**
    - The range must be defined by `DefineRange()`
- **Inputs:** None
- **Actions:** None
- **Outputs:**
    - The sheet number of the last column in the data range - `Long`

### FirstColLetter()
- **Prerequisits:**
    - The range must be defined by `DefineRange()`
    - `FirstColNum()`
- **Inputs:** None
- **Actions:** None
- **Outputs:**
    - The sheet column letter of the first column in the data range - `String`

### LastColLetter()
- **Prerequisits:**
    - The range must be defined by `DefineRange()`
    - `LastColNum()`
- **Inputs:** None
- **Actions:** None
- **Outputs:**
    - The sheet column letter of the last column in the data range - `String`

### Cell_DataRowColName()
- **Prerequisits:**
    - The range must be defined by `DefineRange()` with a header row
    - The named column must exist
    - `cHeaderRange` class
    - `cDataRange` class
- **Inputs:**
    - *DataRow* As `Long`
    - *ColName* AS `String`
- **Actions:** None
- **Outputs:**
    - Cell in the named column in the defined row of data (not sheet row) - `Range`

### Cell_SheetRowColName()
- **Prerequisits:**
    - The range must be defined by `DefineRange()` with a header row
    - The named column must exist
    - `cHeaderRange` class
    - `cDataRange` class
- **Inputs:**
    - *SheetRow* As `Long`
    - *ColName* AS `String`
- **Actions:** None
- **Outputs:**
    - Cell in the named column in the defined sheet row of data - `Range`

### Cell_UnderHeader()
- **Prerequisits:**
    - The range must be defined by `DefineRange()` with a header row
    - The named column must exist
    - `cHeaderRange` class
- **Inputs:**
    - *RowNum* As `Long` (*Must be greater than 0*)
    - *ColName* AS `String`
- **Actions:** None
- **Outputs:**
    - Cell in the named column in the defined number of rows below the header row (does not have to be in the data range) - `Range`

