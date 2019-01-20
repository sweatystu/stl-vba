Option Explicit

' Dependencies
'   -cRange
'   - cColours

' Public Procedures
Sub RemoveRangeFormatting()
    ' Description: Removes all line formatting from a selection
    ' Dependencies: None
    ' Inputs: None
    ' Outputs: None
    Selection.Borders.LineStyle = xlNone
End Sub

Sub FormatRangeAsTable()
    ' Description: Add horizontal lines to the selection, and format the first row as a header
    ' Dependencies:
    '   - cRange
    '   - cColours
    ' Inputs: None
    ' Outputs: None
    Dim c As New cColours
    Dim r As New cRange
    On Error GoTo ErrorHandle:
    r.DefineRange Selection
    With r.rng
        With .Borders(xlInsideHorizontal)
            .LineStyle = xlSolid
            .Weight = xlThin
            .Color = c.Grey.Light
        End With
        With .Borders(xlEdgeBottom)
            .LineStyle = xlSolid
            .Weight = xlMedium
            .Color = c.Grey.Medium
        End With
    End With
    With r.HeaderRange.HeaderRange
        .Font.Bold = True
        With .Borders(xlEdgeBottom)
            .LineStyle = xlSolid
            .Weight = xlMedium
            .Color = c.Grey.Medium
        End With
    End With
    Exit Sub
ErrorHandle:
    custErr.DisplayError "mRangeFormat - FormatRangeAsTable()"
End Sub

Sub FormatRangeVertLines()
    ' Description: Add vertical lines to the selection
    ' Dependencies:
    '   - cColours
    ' Inputs: None
    ' Outputs: None
    Dim c As New cColours
    On Error GoTo ErrorHandle
    With Selection.Borders(xlInsideVertical)
        .LineStyle = xlSolid
        .Weight = xlThin
        .Color = c.Grey.vLight
    End With
    Exit Sub
ErrorHandle:
    custErr.DisplayError "mRangeFormat - FormatRangeVertLines()"
End Sub
