Attribute VB_Name = "modChartStackedBar"
Option Explicit

Private Sub BuildStackedBarChart(Optional ByVal colorMode As String = "FILL")
    Dim cht As Chart

    ActiveSheet.Shapes.AddChart2(-1, xlBarStacked).Select
    ActiveChart.Parent.Duplicate.Select

    Set cht = ActiveChart
    If cht Is Nothing Then Exit Sub

    ' Shared formatting pipeline
    ApplyChartPipeline cht, colorMode

    ' Stacked bar-specific: no tick marks on category axis
    cht.Axes(xlCategory).MajorTickMark = xlTickMarkNone
    cht.Axes(xlCategory).MinorTickMark = xlTickMarkNone

    ' Stacked bar-specific: series layout from config
    cht.ChartGroups(1).Overlap = seriesOverlap
    cht.ChartGroups(1).GapWidth = seriesGapWidth
End Sub

Sub StackedBarChart()
    BuildStackedBarChart
End Sub

Public Sub StackedBar_onAction(control As IRibbonControl)
    BuildStackedBarChart
End Sub

Sub StackedBarChartBlueRamp()
    BuildStackedBarChart "BLUERAMP"
End Sub

Public Sub StackedBarWithBlueRamp_onAction(control As IRibbonControl)
    BuildStackedBarChart "BLUERAMP"
End Sub
