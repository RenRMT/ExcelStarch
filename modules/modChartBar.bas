Attribute VB_Name = "modChartBar"
'==== Module: modChartBar ====
' Clustered and stacked horizontal bar chart variants.
'
' Variants
' --------
'   BarChart        — xlBarClustered: discrete side-by-side bars per category
'   StackedBarChart — xlBarStacked:   series stacked into a single bar per category
'
' Differences
' -----------
'   Chart type:   xlBarClustered vs xlBarStacked
'   RemoveShadow: called for clustered only. Clustered bars can accumulate per-series
'                 shadows from Excel defaults; stacked bars share a single bar body so
'                 shadow removal is not needed.
'
' Everything else is identical: full FILL pipeline, no category-axis tick marks,
' seriesGapWidth from modConfig.
' seriesOverlap: clustered uses modConfig value; stacked is always 100 (slices must be flush).
'
' Note: modChartLollipop wraps BarChart() and post-processes the result into a
' lollipop style. Changes to BuildBarChart may affect lollipop output.
Option Explicit


Private Sub BuildBarChart()
    Dim cht As Chart

    Set cht = GetTargetChart(xlBarClustered)
    If cht Is Nothing Then Exit Sub

    ApplyChartPipeline cht, "FILL"
    Call RemoveShadow(cht)

    cht.Axes(xlCategory).MajorTickMark = xlTickMarkNone
    cht.Axes(xlCategory).MinorTickMark = xlTickMarkNone

    cht.ChartGroups(1).Overlap = seriesOverlap
    cht.ChartGroups(1).GapWidth = seriesGapWidth
End Sub


Private Sub BuildStackedBarChart()
    Dim cht As Chart

    Set cht = GetTargetChart(xlBarStacked)
    If cht Is Nothing Then Exit Sub

    ApplyChartPipeline cht, "FILL"
    Call RemoveShadow(cht)

    cht.Axes(xlCategory).MajorTickMark = xlTickMarkNone
    cht.Axes(xlCategory).MinorTickMark = xlTickMarkNone

    cht.ChartGroups(1).Overlap = 100
    cht.ChartGroups(1).GapWidth = seriesGapWidth
End Sub


Sub BarChart()
    BuildBarChart
End Sub

Sub StackedBarChart()
    BuildStackedBarChart
End Sub
