Attribute VB_Name = "modChartArea"
'==== Module: modChartArea ====
' Area and 100% stacked area chart variants.
'
' Variants
' --------
'   AreaChart    — xlAreaStacked:    series stacked into a single filled area per category
'   Area100Chart — xlAreaStacked100: series stacked to fill 100% per category
'
' Uses the full FILL pipeline. AxisBetweenCategories = False so areas
' fill flush to both chart edges (same pattern as line charts). Tick marks are hidden
' on both axes (consistent with bar/column style). Axis lines are re-formatted to white
' after AxisBetweenCategories assignment, which can re-show them.
Option Explicit


Private Sub BuildStackedAreaChart()
    On Error GoTo CleanFail
    AppFast

    Dim cht As Chart

    Set cht = GetTargetChart(xlAreaStacked)
    If cht Is Nothing Then GoTo CleanExit

    ApplyChartPipeline cht, "FILL", AreaChartDefaults()

    ' Area-specific: axis starts on first data point so areas fill to chart edges
    If cht.HasAxis(xlCategory) Then
        cht.Axes(xlCategory).AxisBetweenCategories = False
        cht.Axes(xlCategory).MajorTickMark = xlTickMarkNone
        cht.Axes(xlCategory).MinorTickMark = xlTickMarkNone
        ' Re-format axis line to white: AxisBetweenCategories assignment can re-show it
        FormatAxisLineWhite cht.Axes(xlCategory)
    End If

    If cht.HasAxis(xlValue) Then
        cht.Axes(xlValue).MajorTickMark = xlTickMarkNone
        cht.Axes(xlValue).MinorTickMark = xlTickMarkNone
        FormatAxisLineWhite cht.Axes(xlValue)
    End If
CleanExit:
    AppRestore
    Exit Sub
CleanFail:
    AppRestore
    MsgError "BuildStackedAreaChart"
End Sub

Private Sub BuildStacked100AreaChart()
    On Error GoTo CleanFail
    AppFast

    Dim cht As Chart

    Set cht = GetTargetChart(xlAreaStacked100)
    If cht Is Nothing Then GoTo CleanExit

    ApplyChartPipeline cht, "FILL", AreaChartDefaults()

    ' Area-specific: axis starts on first data point so areas fill to chart edges
    If cht.HasAxis(xlCategory) Then
        cht.Axes(xlCategory).AxisBetweenCategories = False
        cht.Axes(xlCategory).MajorTickMark = xlTickMarkNone
        cht.Axes(xlCategory).MinorTickMark = xlTickMarkNone
        ' Re-format axis line to white: AxisBetweenCategories assignment can re-show it
        FormatAxisLineWhite cht.Axes(xlCategory)
    End If

    If cht.HasAxis(xlValue) Then
        cht.Axes(xlValue).MajorTickMark = xlTickMarkNone
        cht.Axes(xlValue).MinorTickMark = xlTickMarkNone
        FormatAxisLineWhite cht.Axes(xlValue)
    End If
CleanExit:
    AppRestore
    Exit Sub
CleanFail:
    AppRestore
    MsgError "BuildStacked100AreaChart"
End Sub

Sub AreaChart()
    BuildStackedAreaChart
End Sub

Sub Area100Chart()
    BuildStacked100AreaChart
End Sub
