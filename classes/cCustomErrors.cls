Option Explicit

' Dependencies: None

' Private Constants
Private Const vMarker As String = "!#!marker!#!" ' Marker to be used in error source - used to show that an error is being propogated

' Collection of codes to use in throwing errors
' Start at 513 and increase by 1 for each error type
Private Const vGenericError As Long = 513

' Public Properties
Property Get GenericError() As Long
    ' Description: Defined number to throw in the event of a generic error
    ' Dependencies: None
    ' Inputs: None
    ' Outputs:
    '   - (As Long) A number to denote a generic error
    GenericError = vbObjectError + vGenericError
End Property

' Public Procedures
Sub RaiseError(ByVal procedure As String)
    ' Description: Raises an error, propagating it to the parent procedure, adding the text to the description
    ' Dependencies: None
    ' Inputs:
    '   - (As String)   procedure   - Text describing the location of the error in the script
    '                               - Usually "ClassName - ProcedureName()"
    ' Outputs:
    '   - Raises an error that propagates to the parent procedure
    Dim vSource As String
    If Left(Err.Source, Len(vMarker)) <> vMarker Then
        ' First error raised
        If Not Erl = 0 Then
            vSource = "Line Number: " & Erl & vbNewLine
        End If
        vSource = vMarker & vSource & procedure
    Else
        vSource = Err.Source & vbNewLine & procedure
    End If
    Err.Raise Err.Number, vSource, Err.Description
End Sub

Sub DisplayError(ByVal procedure As String)
    ' Description: Displays the error details in a messagebox
    '               Includes the error type and the location in the script
    ' Dependencies:
    '   - Me.ErrorMessage()
    ' Inputs:
    '   - (As String)   procedure   - Text describing the mother procedure
    '                               - Usually "ModuleName - ProcedureName()"
    ' Outputs:
    '   - Displays a messagebox with error details
    MsgBox Me.ErrorMessage(procedure), VbMsgBoxStyle.vbOKOnly + VbMsgBoxStyle.vbExclamation, "Error"
End Sub

' Public Functions
Function ErrorMessage(ByVal procedure As String) As String
    ' Description: Creates a standard message to display the error
    ' Dependencies: None
    ' Inputs:
    '   - (As String)   procedure   - Text describing the mother procedure
    '                               - Usually "ModuleName - ProcedureName()"
    ' Outputs:
    '   - (As String) The message to be displayed
    Dim msg As String
    msg = "An error occured:"
    msg = msg & vbNewLine & "-- "
    Select Case Err.Number
        Case Me.GenericError
            msg = msg & "Custom Generic Error"
        Case Else
            msg = msg & "Standard Error"
    End Select
    msg = msg & " --" & vbNewLine
    msg = msg & Err.Description
    msg = msg & vbNewLine & vbNewLine
    msg = msg & "Error location is:"
    msg = msg & vbNewLine
    msg = msg & Replace(Err.Source, vMarker, "") & vbNewLine & procedure
    ErrorMessage = msg
End Function
