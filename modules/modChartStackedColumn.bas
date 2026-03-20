Attribute VB_Name = "modChartStackedColumn"
Option Explicit

Private Sub BuildStackedColumnChart()
    Dim cht As Chart

    ActiveSheet.Shapes.AddChart2(-1, xlColumnStacked).Select
    ActiveChart.Parent.Duplicate.Select

    Set cht = ActiveChart
    If cht Is Nothing Then Exit Sub

    ' Shared formatting pipeline
    ApplyChartPipeline cht, "FILL"

    ' Stacked column-specific: no tick marks on category axis
    cht.Axes(xlCategory).MajorTickMark = xlTickMarkNone
    cht.Axes(xlCategory).MinorTickMark = xlTickMarkNone

    ' Stacked column-specific: series layout from config
    cht.ChartGroups(1).Overlap = seriesOverlap
    cht.ChartGroups(1).GapWidth = seriesGapWidth
End Sub

Sub StackedColumnChart()
    BuildStackedColumnChart
End Sub

