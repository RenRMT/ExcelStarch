Attribute VB_Name = "modChartLine"
Option Explicit

Private Sub BuildLineChart()
    Dim cht As Chart

    ActiveSheet.Shapes.AddChart2(-1, xlLine).Select
    ActiveChart.Parent.Duplicate.Select

    Set cht = ActiveChart
    If cht Is Nothing Then Exit Sub

    ' Shared formatting pipeline
    ApplyChartPipeline cht, "LINE"

    ' Line-specific: axis starts on first data point (not between categories)
    cht.Axes(xlCategory).AxisBetweenCategories = False

    ' Line-specific: tick marks outside on both axes
    cht.Axes(xlCategory).MajorTickMark = xlTickMarkOutside
    cht.Axes(xlCategory).MinorTickMark = xlTickMarkNone
    cht.Axes(xlValue).MajorTickMark = xlTickMarkOutside
    cht.Axes(xlValue).MinorTickMark = xlTickMarkNone
End Sub

Sub LineChart()
    BuildLineChart
End Sub
