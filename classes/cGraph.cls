Option Explicit

' Dependencies:
'   - cCustomErrors() class
'   - cColours() class

' Private Variables
Private vChart As ChartObject
Private vColours As cColours

' Public Procedures
Sub DefineGraph(ByRef cht As ChartObject)
    ' Description: Defines the graph to edit. Initialises the colour palette.
    ' Dependencies:
    '   - cColours() class
    ' Inputs:
    '   - (As ChartObject)  cht     - the graph to be edited
    ' Outputs: None
    On Error GoTo ErrorHandle
    Set vChart = cht
    Set vColours = New cColours
    Exit Sub
ErrorHandle:
    custErr.RaiseError "cGraph - DefineGraph()"
End Sub

Sub SetTitleText(ByVal txt As String)
    ' Description: Changes the text of the title
    ' Dependencies:
    '   - Me.Graph()
    ' Inputs:
    '   - (As String)   txt     - the new text of the title
    ' Outputs: None
    On Error GoTo ErrorHandle
    If Not Me.Graph.Chart.HasTitle Then Me.Graph.Chart.HasTitle = True
    Me.Graph.Chart.ChartTitle.Caption = txt
    Exit Sub
ErrorHandle:
    custErr.RaiseError "cGraph - SetTitleText()"
End Sub

Sub SetAxisTitleText(ByVal txt As String, ByVal axType As XlAxisType, Optional ByVal axGroup As XlAxisGroup = XlAxisGroup.xlPrimary)
    ' Description: Changes the text of the defined axis
    ' Dependencies:
    '   - Me.Graph()
    ' Inputs:
    '   - (As String)               txt         - the new text of the title
    '   - (As XlAxisType)           axType      - the axis to change
    '   - (opt. As XlAxis Group)    axGroup     - the axis group (primary / secondary). Defaults to xlPrimary
    ' Outputs: None
    Dim ax As Axis
    On Error GoTo ErrorHandle
    Set ax = Me.Graph.Chart.Axes(axType, axGroup)
    If Not ax.HasTitle Then ax.HasTitle = True
    ax.AxisTitle.Caption = txt
    Exit Sub
ErrorHandle:
    custErr.RaiseError "cGraph - SetAxisTitleText()"
End Sub

Sub RemoveGraphTitle()
    ' Description: Removes the graph title
    ' Dependencies:
    '   - Me.Graph()
    ' Inputs: None
    ' Outputs: None
    On Error GoTo ErrorHandle
    Me.Graph.Chart.HasTitle = False
    Exit Sub
ErrorHandle:
    custErr.RaiseError "cGraph - RemoveGraphTitle()"
End Sub

Sub RemoveAxesTitles()
    ' Description: Removes titles from all axes
    ' Dependencies:
    '   - Me.Graph()
    ' Inputs: None
    ' Outputs: None
    Dim ax As Axis
    On Error GoTo ErrorHandle
    For Each ax In Me.Graph.Chart.Axes
        ax.HasTitle = False
    Next ax
    Exit Sub
ErrorHandle:
    custErr.RaiseError "cGraph - RemoveAxesTitles()"
End Sub

Sub SetGridLines(Optional ByVal axValue As Boolean = False, Optional ByVal axCat As Boolean = False)
    ' Description: Sets whether gridlines are displayed or not against the primary axes
    ' Dependencies:
    '   - Me.Graph()
    ' Inputs:
    '   - (opt. As Boolean) axValue     - Whether the gridlines against the value axis are shown or not. Defaults to false.
    '   - (opt. As Boolean) axCat       - Whether the gridlines against the category axis are shown or not. Defaults to false.
    ' Outputs: None
    Dim ax As Axis
    On Error GoTo ErrorHandle
    For Each ax In Me.Graph.Chart.Axes
        ax.HasMajorGridlines = False
        ax.HasMinorGridlines = False
    Next ax
    Me.Graph.Chart.Axes(XlAxisType.xlValue, XlAxisGroup.xlPrimary).HasMajorGridlines = axValue
    Me.Graph.Chart.Axes(XlAxisType.xlCategory, XlAxisGroup.xlPrimary).HasMajorGridlines = axCat
    Exit Sub
ErrorHandle:
    custErr.RaiseError "cGraph - SetGridLines()"
End Sub

Sub SetLegendPosition(ByVal HasLegend As Boolean, Optional ByVal pos As XlLegendPosition = XlLegendPosition.xlLegendPositionBottom)
    ' Description: Hides or positions the legend
    ' Dependencies:
    '   - Me.Graph()
    ' Inputs:
    '   - (As Boolean)                  HasLegend   - Whether the legend is to be displayed or not
    '   - (opt. As XlLegendPosition)    pos         - Position of the legend, defaults to xlLegendPositionBottom
    On Error GoTo ErrorHandle
    Me.Graph.Chart.HasLegend = HasLegend
    If HasLegend Then Me.Graph.Chart.Legend.Position = pos
    Exit Sub
ErrorHandle:
    custErr.RaiseError "cGraph - SetLegendPosition()"
End Sub

Sub FormatChartGeneral()
    ' Description: Applies basic formatting to chart - white background with no borders
    ' Dependencies:
    '   - Me.Graph()
    ' Inputs: None
    ' Outputs: None
    On Error GoTo ErrorHandle
    With Me.Graph.Chart
        With .ChartArea
            .Line.Visible = msoFalse
            .Fill.BackColor.RGB = vColours.White
        End With
        With .PlotArea
            .Line.Visible = msoFalse
            .Fill.BackColor.RGB = vColours.White
        End With
    End With
    Exit Sub
ErrorHandle:
    custErr.RaiseError "cGraph - FormatChartGeneral()"
End Sub

Sub FormatChartTitle()
    ' Description: Applies formatting to the chart title
    ' Dependencies:
    '   - Me.Graph()
    ' Inputs: None
    ' Outputs: None
    On Error GoTo ErrorHandle
    If Not Me.Graph.Chart.HasTitle Then Exit Sub
    With Me.Graph.Chart.ChartTitle
        With .Format.TextFrame2.TextRange.Font
            .Size = 12
            .Fill.ForeColor.RGB = vColours.Grey.Medium
        End With
        .Position = xlChartElementPositionAutomatic
        .Top = 3
        .Left = 15
    End With
    Exit Sub
ErrorHandle:
    custErr.RaiseError "cGraph - FormatChartTitle()"
End Sub

Sub FormatAxisTitles()
    ' Description: Applies formatting to all axis titles
    ' Dependencies:
    '   - Me.Graph()
    ' Inputs: None
    ' Outputs: None
    Dim ax As Axis
    On Error GoTo ErrorHandle
    For Each ax In Me.Graph.Chart.Axes
        If ax.HasTitle Then
            With ax.AxisTitle.Font
                .Size = 9
                .Color = vColours.Grey.Medium
            End With
        End If
    Next ax
    Exit Sub
ErrorHandle:
    custErr.RaiseError "cGraph - FormatAxisTitles()"
End Sub

Sub FormatLegend()
    ' Description: Applies formatting to the legend
    ' Dependencies:
    '   - Me.Graph()
    ' Inputs: None
    ' Outputs: None
    On Error GoTo ErrorHandle
    If Not Me.Graph.Chart.HasLegend Then Exit Sub
    With Me.Graph.Chart.Legend.Format.TextFrame2.TextRange.Font
        .Size = 9
        .Fill.ForeColor.RGB = vColours.Grey.Medium
    End With
    Exit Sub
ErrorHandle:
    custErr.RaiseError "cGraph - FormatLegend()"
End Sub

Sub FormatAxes()
    ' Description: Applies formatting to each axis
    ' Dependencies:
    '   - Me.Graph()
    ' Inputs: None
    ' Outputs: None
    Dim ax As Axis
    On Error GoTo ErrorHandle
    For Each ax In Me.Graph.Chart.Axes
        ax.Border.Weight = xlThin
        ax.Border.Color = vColours.Grey.Light
        With ax.TickLabels.Font
            .Size = 9
            .Color = vColours.Grey.Medium
        End With
    Next ax
    Exit Sub
ErrorHandle:
    custErr.RaiseError "cGraph - FormatAxes()"
End Sub

Sub FormatGridlines()
    ' Description: Applies formatting to the gridlines
    ' Dependencies:
    '   - Me.Graph()
    ' Inputs: None
    ' Outputs: None
    Dim ax As Axis
    On Error GoTo ErrorHandle
    For Each ax In Me.Graph.Chart.Axes
        If ax.HasMajorGridlines Then
            With ax.MajorGridlines.Border
                .LineStyle = xlSolid
                .Weight = xlHairline
                .Color = vColours.Grey.Light
            End With
        End If
        If ax.HasMinorGridlines Then
            With ax.MinorGridlines.Border
                .LineStyle = xlSolid
                .Weight = xlHairline
                .Color = vColours.Grey.vLight
            End With
        End If
    Next ax
    Exit Sub
ErrorHandle:
    custErr.RaiseError "cGraph - FormatGridlines()"
End Sub

' Public Properties
Property Get Graph() As ChartObject
    'Description: Returns the chart
    ' Dependencies:
    '   - Chart must have been defined previously
    ' Inputs: None
    ' Outputs:
    '   - (As ChartObject) The stored chart
    On Error GoTo ErrorHandle
    If vChart Is Nothing Then Err.Raise custErr.GenericError, Description:="The graph has not been defined"
    Set Graph = vChart
    Exit Property
ErrorHandle:
    custErr.RaiseError "cGraph - Get Graph()"
End Property
