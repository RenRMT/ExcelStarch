Attribute VB_Name = "modChartScatter"
'==== Module: modChartScatter ====
' Scatter and bubble chart variants.
'
' Variants
' --------
'   ScatterChart — xlXYScatter: X/Y scatter plot with markers only (no connecting lines)
'   BubbleChart  — xlBubble:    scatter with bubble size as a third data dimension
'
' Both variants use the full FILL pipeline so marker and bubble fills are coloured
' by the brand palette via Format.Fill.ForeColor.rgb. Tick marks are applied outside
' on both axes (scatter axes are value axes, not category axes). HasAxis guards are
' used for axis operations in case a particular chart variant omits an axis.
' Axis lines are re-hidden after tick-mark assignment.
Option Explicit


Private Sub BuildScatterChart()
    On Error GoTo CleanFail

    Dim cht As Chart

    Set cht = GetTargetChart(xlXYScatter)
    If cht Is Nothing Then Exit Sub

    ApplyChartPipeline cht, "FILL", ScatterChartDefaults()

    ' Scatter-specific: tick marks outside on both axes
    If cht.HasAxis(xlCategory) Then
        cht.Axes(xlCategory).MajorTickMark = xlTickMarkOutside
        cht.Axes(xlCategory).MinorTickMark = xlTickMarkNone
    End If
    If cht.HasAxis(xlValue) Then
        cht.Axes(xlValue).MajorTickMark = xlTickMarkOutside
        cht.Axes(xlValue).MinorTickMark = xlTickMarkNone
    End If

    ' Re-hide axis lines: tick-mark assignment can re-show them
    If cht.HasAxis(xlValue) Then cht.Axes(xlValue).Format.Line.Visible = msoFalse
    If cht.HasAxis(xlCategory) Then
        cht.Axes(xlCategory).Select
        Selection.Format.Line.Visible = msoFalse
    End If
    Exit Sub
CleanFail:
    MsgError "BuildScatterChart"
End Sub


Private Sub BuildBubbleChart()
    On Error GoTo CleanFail

    Dim cht As Chart

    Set cht = GetTargetChart(xlBubble)
    If cht Is Nothing Then Exit Sub

    ApplyChartPipeline cht, "FILL", ScatterChartDefaults()

    ' Bubble-specific: tick marks outside on both axes
    If cht.HasAxis(xlCategory) Then
        cht.Axes(xlCategory).MajorTickMark = xlTickMarkOutside
        cht.Axes(xlCategory).MinorTickMark = xlTickMarkNone
    End If
    If cht.HasAxis(xlValue) Then
        cht.Axes(xlValue).MajorTickMark = xlTickMarkOutside
        cht.Axes(xlValue).MinorTickMark = xlTickMarkNone
    End If

    ' Re-hide axis lines: tick-mark assignment can re-show them
    If cht.HasAxis(xlValue) Then cht.Axes(xlValue).Format.Line.Visible = msoFalse
    If cht.HasAxis(xlCategory) Then
        cht.Axes(xlCategory).Select
        Selection.Format.Line.Visible = msoFalse
    End If
    Exit Sub
CleanFail:
    MsgError "BuildBubbleChart"
End Sub


Sub ScatterChart()
    BuildScatterChart
End Sub

Sub BubbleChart()
    BuildBubbleChart
End Sub
