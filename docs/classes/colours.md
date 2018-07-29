# Colours

- **Filename:** colours.cls
- **Instance Name:** cColours
- **Prerequisits:** None

This class must be called when a defined colour palette is required.

``` VB
Dim vColours As New cColours
```

## General Use
Class provides a defined colour palette. Colours can be chosen using a combination of `c_Hue` and `c_Shade`. Greys can be chosen by providing a grey-scale number (0 - 255) or by using `c_GreyShade` for pre-defined greys.

## Enums
Enums provide a defined list of options to pass to the properties, ensuring that only correctly defined colours are returned.

### c_Hue
A List of colours.
- Red
- Orange
- Yellow
- Green
- Blue
- Purple

*e.g.* `c_Hue.Yellow`

### c_Shade
A list of shades of colours.
- Bold
- Pastel
- Light

*e.g.* `c_Shade.Pastel`

### c_GreyShade
A list of predefined shades of grey.
- vDark
- Dark
- Medium
- Light
- vLight

*e.g.* `c_GreyShade.Medium`

## Properties

### Colour()
- **Prerequisits:** None
- **Inputs:**
    - hue As `c_Hue` - *Colour*
    - shade As `c_Shade` - *Shade of the colour*
- **Actions:**
    - Identify the colour based on the `c_Hue` and `c_Shade` given
- **Outputs:**
    - Number representing a colour

*e.g.* `vColour.Colour(c_Hue.Orange, c_Shade.Light)` will return a light orange colour.

### Grey256()
- **Prerequisits:** None
- **Inputs:**
    - num As Long - *Numeric value between 0 and 255
- **Actions:**
    - Confirm number is between 0 and 255
- **Outputs:**
    - Number representing a grey colour of the value given as an input

*e.g.* `vColour.Grey256(180)` will return a grey colour - RGB(180, 180, 180).

### GreyShade()
- **Prerequisits:** None
- **Inputs:**
    - shade As `c_GreyShade`
- **Actions:**
    - Identify shade of grey based on the `c_GreyShade` given
- **Outputs:**
    - Number representing a shade of grey

## Custom Colours

Custom colours can be added to the class by defining the numeric value of the colour as a *constant* and then building a *Property Get* to access the colour.

To add the custom colour **RGB(255, 67, 109)** (7160831) ...
``` VB
Private Const vCustCol1 As Long = 7160831

Property Get CustCol1() As Long
    CustCol1 = vCustCol1
End Property
```

