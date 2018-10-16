Option Explicit

' Dependencies: None

' Private Constants
Private vLight As Long
Private vPastel As Long
Private vBold As Long

' Public Procedures
Sub Initialise(ByVal l As Long, ByVal p As Long, ByVal b As Long)
    ' Description: Initialises the colour with light, pastel and bold variants
    ' Dependencies: None
    ' Inputs:
    '   - (As Long) l   - The number of the light variant of the colour
    '   - (As Long) p   - The number of the pastel variant of the colour
    '   - (As Long) b   - The number of the bold variant of the colour
    ' Outputs: None
    On Error GoTo ErrorHandle
    vLight = l
    vPastel = p
    vBold = b
    Exit Sub
ErrorHandle:
    custErr.RaiseError "cIndividualColour - Initialise()"
End Sub

' Public Properties
Property Get Light() As Long
    ' Description: Returns the light version of the colour
    ' Dependencies: None
    ' Inputs: None
    ' Outputs:
    '   - (As Long) The colour representing the light version of the colour
    On Error GoTo ErrorHandle
    Light = vLight
    Exit Property
ErrorHandle:
    custErr.RaiseError "cIndividualColour - Get Light()"
End Property

Property Get Pastel() As Long
    ' Description: Returns the pastel version of the colour
    ' Dependencies: None
    ' Inputs: None
    ' Outputs:
    '   - (As Long) The colour representing the pastel version of the colour
    On Error GoTo ErrorHandle
    Pastel = vPastel
    Exit Property
ErrorHandle:
    custErr.RaiseError "cIndividualColour - Get Pastel()"
End Property

Property Get Bold() As Long
    ' Description: Returns the bold version of the colour
    ' Dependencies: None
    ' Inputs: None
    ' Outputs:
    '   - (As Long) The colour representing the bold version of the colour
    On Error GoTo ErrorHandle
    Bold = vBold
    Exit Property
ErrorHandle:
    custErr.RaiseError "cIndividualColour - Get Bold()"
End Property
