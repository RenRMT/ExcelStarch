Attribute VB_Name = "modChartArea"
Option Explicit

Private Sub BuildAreaChart()
    Dim cht As Chart

    ActiveSheet.Shapes.AddChart2(-1, xlAreaStacked).Select
    ActiveChart.Parent.Duplicate.Select

    Set cht = ActiveChart
    If cht Is Nothing Then Exit Sub

    ' Shared formatting pipeline
    ApplyChartPipeline cht, "FILL"

    ' Area-specific: axis starts on first data point (not between categories)
    cht.Axes(xlCategory).AxisBetweenCategories = False

    ' Area-specific: tick marks outside on category axis
    cht.Axes(xlCategory).MajorTickMark = xlTickMarkOutside
    cht.Axes(xlCategory).MinorTickMark = xlTickMarkNone
End Sub

Sub AreaChart()
    BuildAreaChart
End Sub

