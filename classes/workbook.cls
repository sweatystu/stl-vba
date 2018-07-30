Option Explicit

Private vWB As workbook

Property Get wb() As workbook
    On Error GoTo ErrorHandle
    Set wb = vWB
    Exit Property
ErrorHandle:
    custErr.RaiseError "cWorkbook - Get wb()"
End Property


' Create new workbook
Sub NewWorkbook(Optional ByVal NewSheets As Long = 1)
    Dim CurrentSheets As Long
    
    On Error GoTo ErrorHandle
    ' Change sheets in new workbook
    CurrentSheets = Application.SheetsInNewWorkbook
    Application.SheetsInNewWorkbook = NewSheets
    ' Create new workbook
    Set vWB = Workbooks.Add
    ' Return sheets in new workbook
    Application.SheetsInNewWorkbook = CurrentSheets
    Exit Sub
ErrorHandle:
    custErr.RaiseError "cWorkbook - NewWorkbook()"
End Sub

' Open workbook
Sub OpenWorkbook(ByVal wb_location As String, Optional ByVal wb_readonly As Boolean = True, Optional ByVal wb_editable As Boolean = False, Optional ByVal wb_password As String = "")
    On Error GoTo ErrorHandle
    
    ' Ensure editable if read/write
    If Not wb_readonly Then wb_editable = True
    
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


' Close workbook
Sub CloseWorkbook()
    On Error GoTo ErrorHandle
    Me.wb.Close False
    Set vWB = Nothing
    Exit Sub
ErrorHandle:
    custErr.RaiseError "cWorkbook - CloseWorkbook()"
End Sub


' Save workbook and close
Sub SaveCloseWorkbook()
    On Error GoTo ErrorHandle
    Me.wb.Save
    Me.CloseWorkbook
    Set vWB = Nothing
    Exit Sub
ErrorHandle:
    custErr.RaiseError "cWorkbook - SaveCloseWorkbook()"
End Sub
