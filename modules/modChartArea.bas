Attribute VB_Name = "modChartArea"
Option Explicit

Private Sub BuildAreaChart(Optional ByVal colorMode As String = "FILL")
    Dim cht As Chart

    ActiveSheet.Shapes.AddChart2(-1, xlAreaStacked).Select
    ActiveChart.Parent.Duplicate.Select

    Set cht = ActiveChart
    If cht Is Nothing Then Exit Sub

    ' Shared formatting pipeline
    ApplyChartPipeline cht, colorMode

    ' Area-specific: axis starts on first data point (not between categories)
    cht.Axes(xlCategory).AxisBetweenCategories = False

    ' Area-specific: tick marks outside on category axis
    cht.Axes(xlCategory).MajorTickMark = xlTickMarkOutside
    cht.Axes(xlCategory).MinorTickMark = xlTickMarkNone
End Sub

Sub AreaChart()
    BuildAreaChart
End Sub

Public Sub Area_onAction(control As IRibbonControl)
    BuildAreaChart
End Sub

Sub AreaChartBlueRamp()
    BuildAreaChart "BLUERAMP"
End Sub

Public Sub AreaWithBlueRamp_onAction(control As IRibbonControl)
    BuildAreaChart "BLUERAMP"
End Sub
