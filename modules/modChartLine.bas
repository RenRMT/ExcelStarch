Attribute VB_Name = "modChartLine"
Option Explicit

Private Sub BuildLineChart()
    Dim cht As Chart

    Set cht = GetTargetChart(xlLine)
    If cht Is Nothing Then Exit Sub

    ' Shared formatting pipeline
    ApplyChartPipeline cht, "LINE"
    Call RemoveShadow(cht)

    ' Line-specific: axis starts on first data point (not between categories)
    cht.Axes(xlCategory).AxisBetweenCategories = False

    ' Line-specific: tick marks outside on both axes
    cht.Axes(xlCategory).MajorTickMark = xlTickMarkOutside
    cht.Axes(xlCategory).MinorTickMark = xlTickMarkNone
    cht.Axes(xlValue).MajorTickMark = xlTickMarkOutside
    cht.Axes(xlValue).MinorTickMark = xlTickMarkNone

    ' Re-hide axis lines: setting AxisBetweenCategories and tick marks can re-show them.
    cht.Axes(xlValue).Format.Line.Visible = msoFalse
    cht.Axes(xlCategory).Select
    Selection.Format.Line.Visible = msoFalse
End Sub

Sub LineChart()
    BuildLineChart
End Sub
