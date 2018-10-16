Option Explicit

' Dependencies:
'   - cCustomErrors
'   - cIndividualColour
'   - cGrey

' Private Constants
Private vRed As cIndividualColour
Private vOrange As cIndividualColour
Private vYellow As cIndividualColour
Private vGreen As cIndividualColour
Private vBlue As cIndividualColour
Private vPurple As cIndividualColour
Private vGrey As cGrey

' Private Procedures
Private Sub Class_Initialize()
    ' Description: Script run when class is initialised
    '               Sets defined colours
    ' Dependencies:
    '   - cIndividualColour class
    ' Inputs: None
    ' Outputs: None
    On Error GoTo ErrorHandle
    Set vRed = New cIndividualColour
    vRed.Initialise RGB(255, 200, 200), RGB(240, 85, 90), RGB(255, 25, 50)
    Set vOrange = New cIndividualColour
    vOrange.Initialise RGB(255, 225, 210), RGB(250, 180, 100), RGB(255, 130, 30)
    Set vYellow = New cIndividualColour
    vYellow.Initialise RGB(255, 255, 180), RGB(240, 240, 100), RGB(255, 192, 0)
    Set vGreen = New cIndividualColour
    vGreen.Initialise RGB(220, 250, 210), RGB(150, 190, 90), RGB(90, 120, 50)
    Set vBlue = New cIndividualColour
    vBlue.Initialise RGB(220, 230, 240), RGB(130, 180, 240), RGB(0, 112, 191)
    Set vPurple = New cIndividualColour
    vPurple.Initialise RGB(230, 210, 230), RGB(185, 145, 205), RGB(110, 70, 140)
    Set vGrey = New cGrey
    Exit Sub
ErrorHandle:
    custErr.RaiseError "cColours - Class_Initialize()"
End Sub

' Public Properties
Property Get Red() As cIndividualColour
    ' Description: Returns the red colour
    ' Dependencies:
    '   - cIndividualColour class
    '   - cIndividualColour class must have been initialised
    ' Inputs: None
    ' Outputs:
    '   - (As cIndividualColour) Red object
    On Error GoTo ErrorHandle
    Set Red = vRed
    Exit Property
ErrorHandle:
    custErr.RaiseError "cColours - Get Red()"
End Property

Property Get Orange() As cIndividualColour
    ' Description: Returns the orange colour
    ' Dependencies:
    '   - cIndividualColour class
    '   - cIndividualColour class must have been initialised
    ' Inputs: None
    ' Outputs:
    '   - (As cIndividualColour) Orange object
    On Error GoTo ErrorHandle
    Set Orange = vOrange
    Exit Property
ErrorHandle:
    custErr.RaiseError "cColours - Get Orange()"
End Property

Property Get Yellow() As cIndividualColour
    ' Description: Returns the yello colour
    ' Dependencies:
    '   - cIndividualColour class
    '   - cIndividualColour class must have been initialised
    ' Inputs: None
    ' Outputs:
    '   - (As cIndividualColour) Yellow object
    On Error GoTo ErrorHandle
    Set Yellow = vYellow
    Exit Property
ErrorHandle:
    custErr.RaiseError "cColours - Get Yellow()"
End Property

Property Get Green() As cIndividualColour
    ' Description: Returns the green colour
    ' Dependencies:
    '   - cIndividualColour class
    '   - cIndividualColour class must have been initialised
    ' Inputs: None
    ' Outputs:
    '   - (As cIndividualColour) Green object
    On Error GoTo ErrorHandle
    Set Green = vGreen
    Exit Property
ErrorHandle:
    custErr.RaiseError "cColours - Get Green()"
End Property

Property Get Blue() As cIndividualColour
    ' Description: Returns the blue colour
    ' Dependencies:
    '   - cIndividualColour class
    '   - cIndividualColour class must have been initialised
    ' Inputs: None
    ' Outputs:
    '   - (As cIndividualColour) Blue object
    On Error GoTo ErrorHandle
    Set Blue = vBlue
    Exit Property
ErrorHandle:
    custErr.RaiseError "cColours - Get Blue()"
End Property

Property Get Purple() As cIndividualColour
    ' Description: Returns the purple colour
    ' Dependencies:
    '   - cIndividualColour class
    '   - cIndividualColour class must have been initialised
    ' Inputs: None
    ' Outputs:
    '   - (As cIndividualColour) Purple object
    On Error GoTo ErrorHandle
    Set Purple = vPurple
    Exit Property
ErrorHandle:
    custErr.RaiseError "cColours - Get Purple()"
End Property

Property Get Grey() As cGrey
    ' Description: Returns the grey colour
    ' Dependencies:
    '   - cGrey class
    '   - cGrey class must have been initialised
    ' Inputs: None
    ' Outputs:
    '   - (As cGrey) Grey object
    On Error GoTo ErrorHandle
    Set Grey = vGrey
    Exit Property
ErrorHandle:
    custErr.RaiseError "cColours - Get Grey()"
End Property
