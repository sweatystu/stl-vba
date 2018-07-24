Option Explicit

' Marker to be used in error source - used to show that an error is being propogated
Private Const vMarker As String = "!#!marker!#!"

' Collection of codes to use in throwing errors
Private Const vGenericError As Long = 513

Property Get GenericError() As Long
    GenericError = vbObjectError + vGenericError
End Property




Sub RaiseError(ByVal ErrorNum As Long, ByVal Description As String, ByVal src As String, ByVal procedure As String, Optional ByVal LineNum As Long = 0)
    Dim vSource As String
    If Left(src, Len(vMarker)) <> vMarker Then
        ' First error raised
        If LineNum <> 0 Then
            vSource = "Line Number: " & LineNum & vbNewLine
        End If
        vSource = vMarker & vSource & procedure
    Else
        vSource = src & vbNewLine & procedure
    End If

    Err.Raise ErrorNum, vSource, Description
End Sub


Sub DisplayError(ByVal src As String, ByVal desc As String, ByVal procedure As String, ByVal errNum As Long)
    Dim msg As String
    msg = "An error occured:"
    msg = msg & vbNewLine
    Select Case errNum
        Case Me.GenericError
            msg = msg & "Custom Generic Error"
        Case Else
            msg = msg & "Non-Custom Error"
    End Select
    msg = msg & vbNewLine
    msg = msg & desc
    msg = msg & vbNewLine & vbNewLine
    msg = msg & "Error location is:"
    msg = msg & vbNewLine
    msg = msg & Replace(src, vMarker, "") & vbNewLine & procedure
    MsgBox msg, Title:="Error"
End Sub



