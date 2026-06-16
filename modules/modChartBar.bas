Attribute VB_Name = "modChartBar"
'==== Module: modChartBar ====
' Clustered and stacked horizontal bar chart variants.
'
' Variants
' --------
'   BarChart            - xlBarClustered:   discrete side-by-side bars per category
'   StackedBarChart     - xlBarStacked:     series stacked into a single bar per category
'   Stacked100BarChart  - xlBarStacked100:  series stacked to fill 100% per category
'
' Differences
' -----------
'   Chart type:   xlBarClustered vs xlBarStacked vs xlBarStacked100
'   Overlap:      clustered uses the modConfig seriesOverlap value; stacked variants
'                 are always 100 (slices must be flush).
'
' Everything else is identical, so the three variants share one private worker
' (BuildBarFamily); each public entry point passes its chart type, overlap, and an
' error-context breadcrumb. RemoveShadow is called for every variant: clustered
' bars can accumulate per-series shadows from Excel defaults, and it is a no-op
' where there is nothing to remove.
'
' Note: modChartLollipop wraps BarChart() and post-processes the result into a
' lollipop style. Changes to BuildBarFamily may affect lollipop output.
Option Explicit


' Shared builder for all horizontal bar variants. The error breadcrumb is passed
' in so a failure still reports which public variant was invoked.
Private Sub BuildBarFamily(ByVal ct As Long, ByVal overlap As Long, ByVal breadcrumb As String)
    On Error GoTo CleanFail
    AppFast

    Dim cht As Chart

    Set cht = GetTargetChart(ct)
    If cht Is Nothing Then GoTo CleanExit

    ApplyChartPipeline cht, "FILL", BarChartDefaults()
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


Sub BarChart()
    BuildBarFamily xlBarClustered, seriesOverlap, "BuildBarChart"
End Sub

Sub StackedBarChart()
    BuildBarFamily xlBarStacked, 100, "BuildStackedBarChart"
End Sub

Sub Stacked100BarChart()
    BuildBarFamily xlBarStacked100, 100, "BuildStacked100BarChart"
End Sub
