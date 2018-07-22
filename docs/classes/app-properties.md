# App Properties

This *class* is an instance of MS Excel. Used to quickly enforce known consistent behaviour in Excel when macros are run.

This class uses initialisation and termination procedures to automatically apply known behaviours. Because of this, allowing procedures to be terminated with an `End` command must be discouraged. Errors should be caught and the termination procedure allowed to run.

## Constants
- vNoSheets - The number of sheets currently in a new workbook. This is temporarily overwritten with known behaviour.
- vCalcOptions - Whether calculation are currently set to automatic or manual. This is temporarily overwritten with known behaviour.
- vEvents - Whether events are currently listened for or not. This is temporarily overwritten with known behaviour.
- vAlerts - Whether alerts are currently displayed or not. This is temporarily overwritten with known behaviour.

## Public Procedures

### SheetsInNew(ByRef i As Long)
Changes the number of sheets in a new workbook to the number passed as an argument.


## Private Procedures

### Class_Initialize()
Records the current values of various application variables and overwrites them with known behaviour.
- Sets the number of sheets in a new workbook to 1
- Sets *calculations* to *manual*
- Disables *events*
- Disables *alerts*
- Turns off *screen updating*

### Class_Terminate()
Returns the variables defined in `Class_Initialize()` to their original setting.


