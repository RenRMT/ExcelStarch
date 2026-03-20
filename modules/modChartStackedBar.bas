Attribute VB_Name = "modChartStackedBar"
Option Explicit

Private Sub BuildStackedBarChart()
    Dim cht As Chart

    ActiveSheet.Shapes.AddChart2(-1, xlBarStacked).Select
    ActiveChart.Parent.Duplicate.Select

    Set cht = ActiveChart
    If cht Is Nothing Then Exit Sub

    ' Shared formatting pipeline
    ApplyChartPipeline cht, "FILL"

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

