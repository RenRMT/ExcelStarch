Attribute VB_Name = "modChartArea"
'==== Module: modChartArea ====
' Area and 100% stacked area chart variants.
'
' Variants
' --------
'   AreaChart    — xlArea:           areas from a zero baseline, series overlaid
'
' Uses the full FILL pipeline. AxisBetweenCategories = False so areas
' fill flush to both chart edges (same pattern as line charts). Tick marks are hidden
' on both axes (consistent with bar/column style). Axis lines are re-hidden after
' AxisBetweenCategories assignment, which can re-show them.
Option Explicit


Private Sub BuildStackedAreaChart()
    On Error GoTo CleanFail

    Dim cht As Chart

    Set cht = GetTargetChart(xlAreaStacked)
    If cht Is Nothing Then Exit Sub

    ApplyChartPipeline cht, "FILL"

    ' Area-specific: axis starts on first data point so areas fill to chart edges
    If cht.HasAxis(xlCategory) Then
        cht.Axes(xlCategory).AxisBetweenCategories = False
        cht.Axes(xlCategory).MajorTickMark = xlTickMarkNone
        cht.Axes(xlCategory).MinorTickMark = xlTickMarkNone
        ' Re-hide axis line: AxisBetweenCategories assignment can re-show it
        cht.Axes(xlCategory).Select
        Selection.Format.Line.Visible = msoFalse
    End If

    If cht.HasAxis(xlValue) Then
        cht.Axes(xlValue).MajorTickMark = xlTickMarkNone
        cht.Axes(xlValue).MinorTickMark = xlTickMarkNone
        cht.Axes(xlValue).Format.Line.Visible = msoFalse
    End If
    Exit Sub
CleanFail:
    MsgError "BuildStackedAreaChart"
End Sub

Sub AreaChart()
    BuildStackedAreaChart
End Sub
