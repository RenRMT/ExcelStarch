Attribute VB_Name = "modChartArea"
'==== Module: modChartArea ====
' Area and 100% stacked area chart variants.
'
' Variants
' --------
'   AreaChart    — xlAreaStacked:    series stacked into a single filled area per category
'   Area100Chart — xlAreaStacked100: series stacked to fill 100% per category
'
' The two variants differ only in chart type, so they share one private worker
' (BuildAreaFamily); each public entry point passes its chart type and an
' error-context breadcrumb.
'
' Uses the full FILL pipeline. AxisBetweenCategories = False so areas
' fill flush to both chart edges (same pattern as line charts). Tick marks are hidden
' on both axes (consistent with bar/column style). Axis lines are re-formatted to white
' after AxisBetweenCategories assignment, which can re-show them.
Option Explicit


' Shared builder for both stacked-area variants. The error breadcrumb is passed
' in so a failure still reports which public variant was invoked.
Private Sub BuildAreaFamily(ByVal ct As Long, ByVal breadcrumb As String)
    On Error GoTo CleanFail
    AppFast

    Dim cht As Chart

    Set cht = GetTargetChart(ct)
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
    MsgError breadcrumb
End Sub


Sub AreaChart()
    BuildAreaFamily xlAreaStacked, "BuildStackedAreaChart"
End Sub

Sub Area100Chart()
    BuildAreaFamily xlAreaStacked100, "BuildStacked100AreaChart"
End Sub
