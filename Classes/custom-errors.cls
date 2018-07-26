Option Explicit

' Marker to be used in error source - used to show that an error is being propogated
Private Const vMarker As String = "!#!marker!#!"

' Collection of codes to use in throwing errors
Private Const vGenericError As Long = 513

Property Get GenericError() As Long
    GenericError = vbObjectError + vGenericError
End Property




Sub RaiseError(ByVal procedure As String)
    Dim vSource As String
    If Left(Err.Source, Len(vMarker)) <> vMarker Then
        ' First error raised
        If Erl <> 0 Then
            vSource = "Line Number: " & Erl & vbNewLine
        End If
        vSource = vMarker & vSource & procedure
    Else
        vSource = Err.Source & vbNewLine & procedure
    End If

    Err.Raise Err.Number, vSource, Err.Description
End Sub


Sub DisplayError(ByVal procedure As String)
    Dim msg As String
    msg = "An error occured:"
    msg = msg & vbNewLine & "-- "
    Select Case Err.Number
        Case Me.GenericError
            msg = msg & "Custom Generic Error"
        Case Else
            msg = msg & "Non-Custom Error"
    End Select
    msg = msg & " --" & vbNewLine
    msg = msg & Err.Description
    msg = msg & vbNewLine & vbNewLine
    msg = msg & "Error location is:"
    msg = msg & vbNewLine
    msg = msg & Replace(Err.Source, vMarker, "") & vbNewLine & procedure
    MsgBox msg, Title:="Error"
End Sub
