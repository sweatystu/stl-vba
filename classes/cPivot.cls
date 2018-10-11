Option Explicit

' Dependencies:
'   - cCustomErrors

' Private Variables
Private pvt As PivotTable

' Public Set Properties
Property Set Pivot(ByRef pt As PivotTable)
    ' Description: Sets the stored pivot table
    ' Dependencies: None
    ' Inputs:
    '   - (As PivotTable)   pt  - The pivot table to be stored
    ' Outputs: None
    On Error GoTo ErrorHandle
    Set pvt = pt
    Exit Property
ErrorHandle:
    custErr.RaiseError "cPivot - Set Pivot()"
End Property

' Public Get Properties
Property Get Pivot() As PivotTable
    ' Description: Returns the stored pivot table
    ' Dependencies:
    '   - Pivot table must have previously been defined
    ' Inputs: None
    ' Outputs:
    '   - (As PivotTable) The stored pivot-table
    On Error GoTo ErrorHandle
    If pvt Is Nothing Then Err.Raise custErr.GenericError, Description:="The pivot table has not been defined"
    Set Pivot = pvt
    Exit Property
ErrorHandle:
    custErr.RaiseError "cPivot - Get Pivot()"
End Property

Property Get PivotFieldExists(ByVal FieldName As String) As Boolean
    ' Description: Confirms whether a field with the given name exists or not
    ' Dependencies:
    '   - Me.Pivot()
    ' Inputs:
    '   - (As String)   FieldName   - Name of the field to identify in the pivot
    ' Outputs:
    '   -(As Boolean) Whether the name field exists or not
    Dim pvtFld As PivotField
    On Error Resume Next
    Set pvtFld = Me.Pivot.PivotFields(FieldName)
    PivotFieldExists = Err.Number = 0
    Exit Property
End Property

' Public Procedures
Sub RefreshPivot()
    ' Description: Refreshes the data in the pivot table
    ' Dependencies:
    '   - Me.Pivot()
    ' Inputs: None
    ' Outputs: None
    On Error GoTo ErrorHandle
    Me.Pivot.PivotCache.Refresh
    Me.Pivot.RefreshTable
    Exit Sub
ErrorHandle:
    custErr.RaiseError "cPivot - RefreshPivot()"
End Sub

Sub ClearPivotFields()
    ' Description: Clear all fields from the pivot
    '               Removes fields from filters, columns and rows
    ' Dependencies:
    '   - Me.Pivot()
    ' Inputs: None
    ' Outputs: None
    Dim pvtField As PivotField
    On Error GoTo ErrorHandle
    For Each pvtField In Me.Pivot.PivotFields
        pvtField.Orientation = XlPivotFieldOrientation.xlHidden
    Next pvtField
    Exit Sub
ErrorHandle:
    custErr.RaiseError "cPivot - ClearPivotFields()"
End Sub

Sub ClearDataFields()
    ' Description: Clear all values from the pivot
    '               Removes fields from values
    ' Dependencies:
    '   - Me.Pivot()
    ' Inputs: None
    ' Outputs: None
    Dim pvtField As PivotField
    On Error GoTo ErrorHandle
    For Each pvtField In Me.Pivot.DataFields
        pvtField.Orientation = XlPivotFieldOrientation.xlHidden
    Next pvtField
    Exit Sub
ErrorHandle:
    custErr.RaiseError "cPivot - ClearDataFields()"
End Sub

Sub ClearAllFields()
    ' Description: Clears all fields and values from the pivot
    '               Removes all fields from filters, columns, rows and values
    ' Dependencies:
    '   - Me.ClearPivotFields()
    '   - Me.ClearDataFields()
    ' Inputs: None
    ' Outputs: None
    On Error GoTo ErrorHandle
    Me.ClearPivotFields
    Me.ClearDataFields
    Exit Sub
ErrorHandle:
    custErr.RaiseError "cPivot - ClearAllFields()"
End Sub

Sub RemoveAllFilters()
    ' Description: Removes any applied filters
    ' Dependencies:
    '   - Me.Pivot()
    ' Inputs: None
    ' Outputs: None
    Dim pvtField As PivotField
    On Error GoTo ErrorHandle
    For Each pvtField In Me.Pivot.PivotFields
        pvtField.ClearAllFilters
    Next pvtField
    Exit Sub
ErrorHandle:
    custErr.RaiseError "cPivot - RemoveAllFilters()"
End Sub

Sub DefineFieldOrientation(ByVal FieldName As String, Optional ByVal FieldOrientation As XlPivotFieldOrientation = XlPivotFieldOrientation.xlHidden)
    ' Description: Gives the defined pivot field the defined orientation
    ' Dependencies:
    '   - Me.Pivot()
    '   - Me.PivotFieldExists()
    ' Inputs:
    '   - (As String)                           FieldName           - The name of the field to be altered
    '   - (As XlPivotFieldOrientation, opt)     FieldOrientation    - The orientation the field is to be changed to
    '                                                               - filter / column / row / data / hidden
    '                                                               - Default is hidden
    ' Outputs: None
    On Error GoTo ErrorHandle
    If Not Me.PivotFieldExists Then Err.Raise custErr.GenericError, Description:="The defined field, " & FieldName & ", could not be identified in the pivot"
    Me.Pivot.PivotFields(FieldName).Orientation = FieldOrientation
    Exit Sub
ErrorHandle:
    custErr.RaiseError "cPivot - DefineFieldOrientation()"
End Sub
