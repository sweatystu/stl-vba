Option Explicit

' Dependencies:
'   - cGraph

' Constants
Const defx As Long = 400 ' default width of a chart
Const defy As Long = 300 ' default height of a chart

' Private Procedures
Private Sub GraphResize(ByRef g As ChartObject, Optional ByVal x As Long = defx, Optional ByVal y As Long = defy)
    ' Description: Resizes graph to defined dimensions
    ' Dependencies: None
    ' Inputs:
    '   - g (As ChartObject)    The graph to resize
    '   - x (As Long)           The horizontal size of the graph
    '   - y (As Long)           The vertical size of the graph
    ' Outputs: None
    On Error GoTo ErrorHandle
    With g
        .Width = x
        .Height = y
    End With
    Exit Sub
ErrorHandle:
    custErr.RaiseError "mGraphs - pvt GraphResize()"
End Sub

Private Sub GraphFormat(ByRef g As ChartObject)
    ' Description: Formats graph with predefined settings
    ' Dependencies:
    '   - cGraph
    ' Inputs:
    '   - g (As ChartObject)    The graph to format
    ' Outputs: None
    Dim grph As New cGraph
    On Error GoTo ErrorHandle
    grph.DefineGraph g
    With grph
        .FormatChartGeneral
        .FormatAxes
        .FormatAxisTitles
        .FormatChartTitle
        .FormatGridlines
        .FormatLegend
    End With
    Exit Sub
ErrorHandle:
    custErr.RaiseError "mGraphs - pvt GraphFormat()"
End Sub


' Public Procedures
Sub FormatGraph()
    ' Description: Apply generic formatting to a graph
    ' Dependencies:
    '   - GraphFormat()
    ' Inputs: None
    ' Outputs: None
    On Error GoTo ErrorHandle
    If ActiveChart Is Nothing Then Err.Raise custErr.GenericError, Description:="No graph selected"
    GraphFormat ActiveChart.Parent
    Exit Sub
ErrorHandle:
    custErr.DisplayError "mGraphs - FormatGraph()"
End Sub

Sub FormatAllGraphs()
    ' Description: Apply generic formatting to all graphs in sheet
    ' Dependencies:
    '   - GraphFormat()
    ' Inputs: None
    ' Outputs: None
    Dim c As ChartObject
    Dim ws As Worksheet
    On Error GoTo ErrorHandle
    Set ws = ActiveSheet
    For Each c In ws.ChartObjects
        GraphFormat c
    Next c
    Exit Sub
ErrorHandle:
    custErr.DisplayError "mGraphs - FormatAllGraphs()"
End Sub

Sub ResizeGraph()
    ' Description: Resize the selected Graph
    ' Dependencies:
    '   - GraphResize()
    '   - A graph must be selected
    ' Inputs: None
    ' Outputs: None
    Dim x As Long
    Dim y As Long
    On Error GoTo ErrorHandle
    If ActiveChart Is Nothing Then Err.Raise custErr.GenericError, Description:="No graph selected"
    x = CLng(InputBox("How wide should the graph be?", "Graph Width", defx))
    y = CLng(InputBox("How tall should the graph be?", "Graph Height", defy))
    GraphResize ActiveChart.Parent, x, y
    Exit Sub
ErrorHandle:
    custErr.DisplayError "mGraphs - ResizeGraph()"
End Sub

Sub ResizeAllGraphs()
    ' Description: Resize the selected Graph
    ' Dependencies:
    '   - GraphResize()
    ' Inputs: None
    ' Outputs: None
    Dim ws As Worksheet
    Dim g As ChartObject
    Dim x As Long
    Dim y As Long
    On Error GoTo ErrorHandle
    Set ws = ActiveSheet
    x = CLng(InputBox("How wide should the graphs be?", "Graph Width", defx))
    y = CLng(InputBox("How tall should the graphs be?", "Graph Height", defy))
    For Each g In ws.ChartObjects
        GraphResize g, x, y
    Next g
    Exit Sub
ErrorHandle:
    custErr.DisplayError "mGraphs - ResizeAllGraphs()"
End Sub
