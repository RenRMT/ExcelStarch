Attribute VB_Name = "modChartColumn"
'==== Module: modChartColumn ====
' Clustered and stacked vertical column chart variants.
'
' Variants
' --------
'   ColumnChart        — xlColumnClustered: discrete side-by-side columns per category
'   StackedColumnChart — xlColumnStacked:   series stacked into a single bar per category
'
' Differences
' -----------
'   Chart type:   xlColumnClustered vs xlColumnStacked
'   RemoveShadow: called for clustered only. Clustered columns can accumulate per-series
'                 shadows from Excel defaults; stacked columns share a single bar body so
'                 shadow removal is not needed.
'
' Everything else is identical: full FILL pipeline, no category-axis tick marks,
' seriesOverlap and seriesGapWidth from modConfig.
Option Explicit


Private Sub BuildColumnChart()
    Dim cht As Chart

    ActiveSheet.Shapes.AddChart2(-1, xlColumnClustered).Select
    ActiveChart.Parent.Duplicate.Select

    Set cht = ActiveChart
    If cht Is Nothing Then Exit Sub

    ApplyChartPipeline cht, "FILL"
    Call RemoveShadow(cht)

    cht.Axes(xlCategory).MajorTickMark = xlTickMarkNone
    cht.Axes(xlCategory).MinorTickMark = xlTickMarkNone

    cht.ChartGroups(1).Overlap = seriesOverlap
    cht.ChartGroups(1).GapWidth = seriesGapWidth
End Sub


Private Sub BuildStackedColumnChart()
    Dim cht As Chart

    ActiveSheet.Shapes.AddChart2(-1, xlColumnStacked).Select
    ActiveChart.Parent.Duplicate.Select

    Set cht = ActiveChart
    If cht Is Nothing Then Exit Sub

    ApplyChartPipeline cht, "FILL"
    Call RemoveShadow(cht)

    cht.Axes(xlCategory).MajorTickMark = xlTickMarkNone
    cht.Axes(xlCategory).MinorTickMark = xlTickMarkNone

    cht.ChartGroups(1).Overlap = seriesOverlap
    cht.ChartGroups(1).GapWidth = seriesGapWidth
End Sub


Sub ColumnChart()
    BuildColumnChart
End Sub

Sub StackedColumnChart()
    BuildStackedColumnChart
End Sub
