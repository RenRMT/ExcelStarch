Attribute VB_Name = "modChartColumn"
'==== Module: modChartColumn ====
' Clustered and stacked vertical column chart variants.
'
' Variants
' --------
'   ColumnChart            — xlColumnClustered:   discrete side-by-side columns per category
'   StackedColumnChart     — xlColumnStacked:     series stacked into a single bar per category
'   Stacked100ColumnChart  — xlColumnStacked100:  series stacked to fill 100% per category
'
' Differences
' -----------
'   Chart type:   xlColumnClustered vs xlColumnStacked vs xlColumnStacked100
'   Overlap:      clustered uses the modConfig seriesOverlap value; stacked variants
'                 are always 100 (slices must be flush).
'
' Everything else is identical, so the three variants share one private worker
' (BuildColumnFamily); each public entry point passes its chart type, overlap, and
' an error-context breadcrumb. RemoveShadow is called for every variant: clustered
' columns can accumulate per-series shadows from Excel defaults, and it is a no-op
' where there is nothing to remove.
Option Explicit


' Shared builder for all vertical column variants. The error breadcrumb is passed
' in so a failure still reports which public variant was invoked.
Private Sub BuildColumnFamily(ByVal ct As Long, ByVal overlap As Long, ByVal breadcrumb As String)
    On Error GoTo CleanFail
    AppFast

    Dim cht As Chart

    Set cht = GetTargetChart(ct)
    If cht Is Nothing Then GoTo CleanExit

    ApplyChartPipeline cht, "FILL", ColumnChartDefaults()
    Call RemoveShadow(cht)

    If cht.HasAxis(xlCategory) Then
        cht.Axes(xlCategory).MajorTickMark = xlTickMarkNone
        cht.Axes(xlCategory).MinorTickMark = xlTickMarkNone
    End If

    cht.ChartGroups(1).Overlap = overlap
    cht.ChartGroups(1).GapWidth = seriesGapWidth
CleanExit:
    AppRestore
    Exit Sub
CleanFail:
    AppRestore
    MsgError breadcrumb
End Sub


Sub ColumnChart()
    BuildColumnFamily xlColumnClustered, seriesOverlap, "BuildColumnChart"
End Sub

Sub StackedColumnChart()
    BuildColumnFamily xlColumnStacked, 100, "BuildStackedColumnChart"
End Sub

Sub Stacked100ColumnChart()
    BuildColumnFamily xlColumnStacked100, 100, "BuildStacked100ColumnChart"
End Sub
