Option Explicit

' Enums
Public Enum c_Hue
    Red
    Orange
    Yellow
    Green
    Blue
    Purple
End Enum

Public Enum c_Shade
    Bold
    Pastel
    Light
End Enum

Public Enum c_GreyShade
    vLight
    Light
    Medium
    Dark
    vDark
End Enum

' Colours
' Convert RGB values to a numeric constant

' Red
Private Const vRedBold As Long = 3283455 ' RGB(255, 25, 50)
Private Const vRedPastel As Long = 5920240 ' RGB(240, 85, 90)
Private Const vRedLight As Long = 13158655 ' RGB(255, 200, 200)
' Orange
Private Const vOrangeBold As Long = 1999615 ' RGB(255, 130, 30)
Private Const vOrangePastel As Long = 6599930 ' RGB(250, 180, 100)
Private Const vOrangeLight As Long = 13820415 ' RGB(255, 225, 210)
' Yellow
Private Const vYellowBold As Long = 49407 ' RGB(255, 192, 0)
Private Const vYellowPastel As Long = 6615280 ' RGB(240, 240, 100)
Private Const vYellowLight As Long = 11862015 ' RGB(255, 255, 180)
' Green
Private Const vGreenBold As Long = 3307610 ' RGB(90, 120, 50)
Private Const vGreenPastel As Long = 5947030 ' RGB(150, 190, 90)
Private Const vGreenLight As Long = 13826780 ' RGB(220, 250, 210)
' Blue
Private Const vBlueBold As Long = 12546048 ' RGB(0, 112, 191)
Private Const vBluePastel As Long = 15774850 ' RGB(130, 180, 240)
Private Const vBlueLight As Long = 15787740 ' RGB(220, 230, 240)
' Purple
Private Const vPurpleBold As Long = 9193070 ' RGB(110, 70, 140)
Private Const vPurplePastel As Long = 13472185 ' RGB(185, 145, 205)
Private Const vPurpleLight As Long = 15127270 ' RGB(230, 210, 230)


' Properties
Property Get Colour(ByVal hue As c_Hue, ByVal shade As c_Shade) As Long
    Dim result As Long
    On Error GoTo ErrorHandle
    
    result = -999 ' set to impossible number
    
    ' Identify colour
    Select Case hue
        Case c_Hue.Red
            Select Case shade
                Case c_Shade.Bold
                    result = vRedBold
                Case c_Shade.Pastel
                    result = vRedPastel
                Case c_Shade.Light
                    result = vRedLight
            End Select
        Case c_Hue.Orange
            Select Case shade
                Case c_Shade.Bold
                    result = vOrangeBold
                Case c_Shade.Pastel
                    result = vOrangePastel
                Case c_Shade.Light
                    result = vOrangeLight
            End Select
        Case c_Hue.Yellow
            Select Case shade
                Case c_Shade.Bold
                    result = vYellowBold
                Case c_Shade.Pastel
                    result = vYellowPastel
                Case c_Shade.Light
                    result = vYellowLight
            End Select
        Case c_Hue.Green
            Select Case shade
                 Case c_Shade.Bold
                    result = vGreenBold
                Case c_Shade.Pastel
                    result = vGreenPastel
                Case c_Shade.Light
                    result = vGreenLight
            End Select
        Case c_Hue.Blue
            Select Case shade
                 Case c_Shade.Bold
                    result = vBlueBold
                Case c_Shade.Pastel
                    result = vBluePastel
                Case c_Shade.Light
                    result = vBlueLight
            End Select
        Case c_Hue.Purple
            Select Case shade
                 Case c_Shade.Bold
                    result = vPurpleBold
                Case c_Shade.Pastel
                    result = vPurplePastel
                Case c_Shade.Light
                    result = vPurpleLight
            End Select
    End Select
    
    ' Confirm colour was found and return result
    If result = -999 Then
        Err.Raise custErr.GenericError, Description:="The colour defined was not recognised"
    Else
        Colour = result
    End If
    Exit Property
    
ErrorHandle:
    custErr.RaiseError "cColours - Colour()"
End Property


Property Get Grey256(ByVal num As Long)
    On Error GoTo ErrorHandle
    
    ' Confirm shade of grey is valid
    If num > 255 Or num < 0 Then
        Err.Raise custErr.GenericError, Description:="The value " & num & " is not between 0 and 255"
    End If
    ' Return result
    Grey256 = RGB(num, num, num)
    Exit Property
    
ErrorHandle:
    custErr.RaiseError "cColours - Grey256()"
End Property


Property Get GreyShade(ByVal shade As c_GreyShade)
    On Error GoTo ErrorHandle
    
    Dim result As Long
    ' Identify shade of grey
    Select Case shade
        Case c_GreyShade.vDark
            result = Me.Grey256(40)
        Case c_GreyShade.Dark
            result = Me.Grey256(80)
        Case c_GreyShade.Medium
            result = Me.Grey256(160)
        Case c_GreyShade.Light
            result = Me.Grey256(200)
        Case c_GreyShade.vLight
            result = Me.Grey256(240)
        Case Else
            Err.Raise custErr.GenericError, Description:="Defined shade of grey not recognised"
    End Select
    ' Return result
    GreyShade = result
    Exit Property
    
ErrorHandle:
    custErr.RaiseError "cColours - GreyShade()"
End Property

