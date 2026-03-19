Attribute VB_Name = "modChartBar"
Option Explicit

Private Sub BuildBarChart()
    Dim cht As Chart

    ActiveSheet.Shapes.AddChart2(-1, xlBarClustered).Select
    ActiveChart.Parent.Duplicate.Select

    Set cht = ActiveChart
    If cht Is Nothing Then Exit Sub

    ' Shared formatting pipeline
    ApplyChartPipeline cht, "FILL"

    ' Bar-specific: remove series shadows
    Call RemoveShadow(cht)

    ' Bar-specific: no tick marks on category axis
    cht.Axes(xlCategory).MajorTickMark = xlTickMarkNone
    cht.Axes(xlCategory).MinorTickMark = xlTickMarkNone

    ' Bar-specific: series layout from config
    cht.ChartGroups(1).Overlap = seriesOverlap
    cht.ChartGroups(1).GapWidth = seriesGapWidth
End Sub

Sub BarChart()
    BuildBarChart
End Sub

Public Sub Bar_onAction(control As IRibbonControl)
    BuildBarChart
End Sub
