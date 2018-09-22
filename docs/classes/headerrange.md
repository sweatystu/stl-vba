# Header Range

- **Filename:** headerrange.cls
- **Instance Name:** cHeaderRange
- **Pre-requisits:**
    - custom-errors.cls - `cCustomErrors` (implemented through *m-global.bas*)

This class is used to extend the `cRange` class. It is to provide details about the header row of a range.

## General Use
This class is automatically initialised as part of the `cRange` class. It should not be used in isolation.

## Properties

### HeaderRange()

#### Set
- **Prerequisits:** None
- **Inputs:**
    - rng As Range - *The range to be used as a range header
- **Actions:** None
- **Outputs:** None

#### Get
- **Prerequisits:**
    - The range must have previously been set
- **Inputs:** None
- **Actions:** None
- **Outputs:**
    - The range of the header - `Range`

### SheetRow()
- **Prerequisits:** None
- **Inputs:** None
- **Actions:**
    - Retrieve sheet row number from range - `HeaderRange()`
- **Outputs:**
    - The sheet row number that the header is in - `long`







