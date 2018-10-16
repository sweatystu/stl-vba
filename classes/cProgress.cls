Option Explicit

' Dependencies:
'   - ProgressForm User Form
'   - cCustomErrors() class
'   - cColours() class

' Private Variables
Private progress_list As MSForms.ListBox
Private progress_text As MSForms.Label
Private errors As MSForms.Label
Private errors_page As MSForms.Page
Private multipage As MSForms.multipage
Private ok_button As MSForms.CommandButton

' Public Procedures
Sub LoadForm(ByVal Title As String, ByVal Description As String)
    ' Description: Loads the progress form and sets the title and description
    ' Dependencies:
    '   - ProgressForm User Form
    ' Inputs:
    '   - (As String)   Title       - The title of the macro being run
    '   - (As String)   Description - A description of the macro
    ' Outputs: None
    On Error GoTo ErrorHandle
    Load ProgressForm
    ProgressForm.lbl_title = Title
    ProgressForm.lbl_description = Description
    Set multipage = ProgressForm.multipage
    Set progress_list = multipage.page_progress.lstbx_progress
    Set progress_text = ProgressForm.lbl_outcome
    Set errors_page = multipage.page_errors
    Set errors = errors_page.lbl_errors
    Set ok_button = ProgressForm.btn_ok
    ProgressForm.Show vbModeless
    DoEvents
    Exit Sub
ErrorHandle:
    custErr.RaiseError "cProgress - LoadForm()"
End Sub

Sub AddTask(ByVal TaskName As String)
    ' Description: Adds a task to the progress form
    ' Dependencies:
    '   - ProgressForm User Form
    ' Inputs:
    '   - (As String)   TaskName    - Name of the task to be completed
    ' Outputs: None
    On Error GoTo ErrorHandle
    DoEvents
    progress_list.AddItem Time & " - " & TaskName & "...", 0
    progress_list.ListIndex = 0
    DoEvents
    Exit Sub
ErrorHandle:
    custErr.RaiseError "cProgress - AddTask()"
End Sub

Sub CompleteTask()
    ' Description: Marks the current task as complete
    ' Dependencies:
    '   - ProgressForm User Form
    ' Inputs: None
    ' Outputs: None
    On Error GoTo ErrorHandle
    DoEvents
    progress_list.List(0) = progress_list.List(0) & "Done"
    DoEvents
    Exit Sub
ErrorHandle:
    custErr.RaiseError "cProgress - CompleteTask()"
End Sub

Sub FailTask()
    ' Description: Marks the current task as failed
    ' Dependencies:
    '   - ProgressForm User Form
    ' Inputs: None
    ' Outputs: None
    On Error GoTo ErrorHandle
    DoEvents
    progress_list.List(0) = progress_list.List(0) & "Fail"
    DoEvents
    Exit Sub
ErrorHandle:
    custErr.RaiseError "cProgress - FailTask()"
End Sub

Sub AddInfo(ByVal info As String)
    ' Description: Adds an information line below the task line
    ' Dependencies:
    '   - ProgressForm User Form
    ' Inputs:
    ' (As String)   info    - Information to be added to the form
    ' Outputs: None
    On Error GoTo ErrorHandle
    DoEvents
    progress_list.AddItem info, 1
    progress_list.ListIndex = 0
    DoEvents
    Exit Sub
ErrorHandle:
    custErr.RaiseError "cProgress - AddInfo()"
End Sub

Sub CompleteMacro()
    ' Description: Sets the form for the end of the macro
    '               Makes OK button enabled and updates output label
    ' Dependencies:
    '   - ProgressForm User Form
    '   - cColours class
    ' Inputs: None
    ' Outputs: None
    Dim Cols As cColours
    On Error GoTo ErrorHandle
    DoEvents
    Set Cols = New cColours
    ProgressForm.Hide
    ok_button.Enabled = True
    progress_text.Caption = "Macro complete!"
    progress_text.BackColor = Cols.Blue.Light
    ProgressForm.Show vbModal
    DoEvents
    Unload ProgressForm
    Exit Sub
ErrorHandle:
    custErr.RaiseError "cProgress - CompleteMacro()"
End Sub

Sub DisplayError(ByVal procedure As String)
    ' Description: Sets the form to display an error
    '               Fails the current task, enables the OK button and displays errors
    ' Dependencies:
    '   - ProgressForm User Form
    '   - cColours class
    ' Inputs:
    '   - (As String)   procedure   - The mother procedure displaying the error
    ' Outputs: None
    Dim Cols As cColours
    errors.Caption = custErr.ErrorMessage(procedure)
    On Error GoTo ErrorHandle
    DoEvents
    Set Cols = New cColours
    ProgressForm.Hide
    Me.FailTask
    ok_button.Enabled = True
    errors_page.Enabled = True
    multipage.Value = 1
    progress_text.Caption = "Error raised!"
    progress_text.BackColor = Cols.Red.Light
    ProgressForm.Show vbModal
    DoEvents
    Unload ProgressForm
    Exit Sub
ErrorHandle:
    custErr.RaiseError "cProgress - DisplayError()"
End Sub
