Attribute VB_Name = "modChartColumn"
Option Explicit

Private Sub BuildColumnChart()
    Dim cht As Chart

    ActiveSheet.Shapes.AddChart2(-1, xlColumnClustered).Select
    ActiveChart.Parent.Duplicate.Select

    Set cht = ActiveChart
    If cht Is Nothing Then Exit Sub

    ' Shared formatting pipeline
    ApplyChartPipeline cht, "FILL"

    ' Column-specific: remove series shadows
    Call RemoveShadow(cht)

    ' Column-specific: no tick marks on category axis
    cht.Axes(xlCategory).MajorTickMark = xlTickMarkNone
    cht.Axes(xlCategory).MinorTickMark = xlTickMarkNone

    ' Column-specific: series layout from config
    cht.ChartGroups(1).Overlap = seriesOverlap
    cht.ChartGroups(1).GapWidth = seriesGapWidth
End Sub

Sub ColumnChart()
    BuildColumnChart
End Sub

