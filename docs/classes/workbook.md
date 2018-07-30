# Workbook

- **Filename:** workbook.cls
- **Instance Name:** cWorkbook
- **Prerequisits:** None

This class is to be used as an extension of the workbook class. It can be used to interact with workbooks in a safe and defined way.

## General Use


## Properties

### wb()
- **Prerequisits:**
    - Workbook must have been initiated by adding a new workbook or by opening an existing workbook
- **Inputs:** None
- **Actions:** None
- **Outputs:**
    - Workbook stored in the class

## Public Procedures
These procedures can be called by procedures in other modules.

### NewWorkbook()
- **Prerequisits:** None
- **Inputs:**
    - *Optional* New Sheets As Long - *number of sheets in new workbook, defaulted to 1*
- **Actions:**
    - Change number of sheets in new workbook to defined number
    - Add new workbook - accessed through `wb` property
    - Return number of sheets in new workbook to previous value
- **Outputs:** None

### OpenWorkbook()
- **Prerequisits:** None
- **Inputs:**
    - location As String - *filepath of the workbook to open*
    - *Optional* wb_readonly As Boolean - *Whether the file is read only or not, defaulted to True*
    - *Optional* wb_editable As Boolean - *Whether the file is editable or not, defaulted to False*
    - *Optional* wb_password As String - *Password to open the file, defaulted to "" (no password)*
- **Actions:**
    - Ensure workbook will be editable if read/write permissions
    - Open workbook according to defined settings - accessed through `wb` property
- **Outputs:** None

### CloseWorkbook()
- **Prerequisits:**
    - Workbook must have been initiated by adding a new workbook or by opening an existing workbook
- **Inputs:** None
- **Actions:**
    - Close workbook without saving
- **Outputs:** None

### SaveCloseWorkbook()
- **Prerequisits:**
    - Workbook must have been initiated by adding a new workbook or by opening an existing workbook
- **Inputs:** None
- **Actions:**
    - Save workbook
    - `CloseWorkbook()` - Close workbook without saving
- **Outputs:** None

## Private Procedures
These procedures can only be called by procedures within the workbook class.

