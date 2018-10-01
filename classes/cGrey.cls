Option Explicit

' Dependencies: None

Property Get Grey256(ByVal vl As Long) As Long
    ' Description: Returns the grey with equal RGB values as given
    ' Dependencies: None
    ' Inputs:
    '   - (As Long)     vl  - The number to use as RGB values
    '                       - Must be between 0 and 255
    ' Outputs:
    '   - (As Long) The number corresponding to the grey with RGB value given
    On Error GoTo ErrorHandle
    If vl < 0 Or vl > 255 Then Err.Raise custErr.GenericError, Description:="The grey value must be between 0 and 255"
    Grey256 = RGB(vl, vl, vl)
    Exit Property
ErrorHandle:
    custErr.RaiseError "cGrey - Get Grey256()"
End Property

Property Get vLight() As Long
    ' Description: Returns the grey defined as very light grey
    ' Dependencies: None
    ' Inputs: None
    ' Outputs:
    '   - (As Long) The number corresponding to the RGB value for very light grey
    On Error GoTo ErrorHandle
    vLight = Me.Grey256(240)
    Exit Property
ErrorHandle:
    custErr.RaiseError "cGrey - Get vLight()"
End Property

Property Get Light() As Long
    ' Description: Returns the grey defined as light grey
    ' Dependencies: None
    ' Inputs: None
    ' Outputs:
    '   - (As Long) The number corresponding to the RGB value for light grey
    On Error GoTo ErrorHandle
    Light = Me.Grey256(200)
    Exit Property
ErrorHandle:
    custErr.RaiseError "cGrey - Get Light()"
End Property

Property Get Medium() As Long
    ' Description: Returns the grey defined as medium grey
    ' Dependencies: None
    ' Inputs: None
    ' Outputs:
    '   - (As Long) The number corresponding to the RGB value for medium grey
    On Error GoTo ErrorHandle
    vLight = Me.Grey256(160)
    Exit Property
ErrorHandle:
    custErr.RaiseError "cGrey - Get Medium()"
End Property

Property Get Dark() As Long
    ' Description: Returns the grey defined as dark grey
    ' Dependencies: None
    ' Inputs: None
    ' Outputs:
    '   - (As Long) The number corresponding to the RGB value for dark grey
    On Error GoTo ErrorHandle
    vLight = Me.Grey256(80)
    Exit Property
ErrorHandle:
    custErr.RaiseError "cGrey - Get Dark()"
End Property

Property Get vDark() As Long
    ' Description: Returns the grey defined as very dark grey
    ' Dependencies: None
    ' Inputs: None
    ' Outputs:
    '   - (As Long) The number corresponding to the RGB value for very dark grey
    On Error GoTo ErrorHandle
    vLight = Me.Grey256(40)
    Exit Property
ErrorHandle:
    custErr.RaiseError "cGrey - Get vDark()"
End Property
