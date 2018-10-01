Option Explicit

' Dependencies: None

' Private Variables
Private vWB As Workbook

' Public Procedures
Sub NewWorkbook()
    ' Description: Creates a new workbook, defined by the class
    ' Dependencies: None
    ' Inputs: None
    ' Outputs: None
    On Error GoTo ErrorHandle
    Set vWB = Workbooks.Add
    Exit Sub
ErrorHandle:
    custErr.RaiseError "cWorkbook - NewWorkbook()"
End Sub

Sub OpenWorkbook(ByVal wb_location As String, Optional ByVal wb_readonly As Boolean = True, Optional ByVal wb_editable As Boolean = False, Optional ByVal wb_password As String = "")
    ' Description: Opens an existing workbook, defined by the class
    ' Dependencies: None
    ' Inputs:
    '   - (As String)               wb_location     - The path and name of the workbook to open
    '   - (As Boolean, optional)    wb_readonly     - Whether the workbook should be opened as read only or not. Defaults to True
    '   - (As Boolean, optional)    wb_editable     - Whether the workbook should be editable or not. Defaults to False
    '   - (As String, optional)     wb_password     - Password required to open the workbook. Defaults to "" with length 0
    ' Outputs: None
    On Error GoTo ErrorHandle
    If Not wb_readonly Then wb_editable = True ' Ensure editable if read/write
    If Len(password) = 0 Then
        ' Open without a password
        Set vWB = Workbooks.Open(Filename:=location, UpdateLinks:=False, Local:=True, readonly:=wb_readonly, editable:=wb_editable)
    Else
        ' Open with a password
        Set vWB = Workbooks.Open(Filename:=location, UpdateLinks:=False, Local:=True, readonly:=wb_readonly, editable:=wb_editable, password:=wb_password)
    End If
    Exit Sub
ErrorHandle:
    custErr.RaiseError "cWorkbook - OpenWorkbook()"
End Sub

Sub CloseWorkbook()
    ' Description: Closes the workbook without saving
    ' Dependencies:
    '   - Workbook must have been created or opened first
    ' Inputs: None
    ' Outputs: None
    On Error GoTo ErrorHandle
    If vWB Is Nothing Then Err.Raise custErr.GenericError, Description:="A workbook has not been defined"
    Me.wb.Close False
    Set vWB = Nothing
    Exit Sub
ErrorHandle:
    custErr.RaiseError "cWorkbook - CloseWorkbook()"
End Sub

Sub SaveCloseWorkbook()
    ' Description: Closes the workbook, saving it first
    ' Dependencies:
    '   - Workbook must have been opened first
    ' Inputs: None
    ' Outputs: None
    On Error GoTo ErrorHandle
    If vWB Is Nothing Then Err.Raise custErr.GenericError, Description:="A workbook has not been defined"
    Me.wb.Save
    Me.CloseWorkbook
    Set vWB = Nothing
    Exit Sub
ErrorHandle:
    custErr.RaiseError "cWorkbook - SaveCloseWorkbook()"
End Sub

' Public Properties
Property Get wb() As Workbook
    ' Description: Returns the defined workbook as an object
    ' Dependencies:
    '   - Workbook must have been created or opened first
    ' Inputs: None
    ' Outputs:
    '   - (As Workbook) The defined workbook
    On Error GoTo ErrorHandle
    Set wb = vWB
    Exit Property
ErrorHandle:
    custErr.RaiseError "cWorkbook - Get wb()"
End Property
