# Range

- **Filename:** range.cls
- **Instance Name:** cRange
- **Prerequisits:** None

This class is to be used as an extension of the range class. It can be used to interact with ranges in a safe, defined and more flexible way.

## General Use
This class must be initialised to be used.

``` VB
Dim rng As New cRange
rng.DefineRange Range("A1:C3"), True, 1
```

This class allows a range to be used with or without a header row. If a header row is used, only the rows *below* the header are considered to be data. This may be useful when importing sheets of data that have metadata in the first few rows.

## Properties

### rng()
- **Prerequisits:**
    - The range must have previously been defined
- **Inputs:** None
- **Actions:** None
- **Outputs:**
    - A defined range of cells

## Public Procedures
These procedures can be accessed by any module with the class initalised.

### DefineRange
- **Prerequisits:** None
- **Inputs:**
    - CellRange As Range - *the range of cells to be stored*
    - *Optional* HasHeaderRow As Boolean - *whether the range has a dedicated header row or not, default is True*
    - *Optional* HeaderRow As Long - *row number, of the range, that the header row is in, defaults to 1*
        - If the defined range starts at row 4 with the header row in the first row, pass 1 as the argument, not 4
- **Actions:**
    - Record range and settings
- **Outputs:** None

## Private Procedures
These procedures can only be accessed by procedures within the range class.



