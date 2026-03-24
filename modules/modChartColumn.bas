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
' seriesGapWidth from modConfig.
' seriesOverlap: clustered uses modConfig value; stacked is always 100 (slices must be flush).
Option Explicit


Private Sub BuildColumnChart()
    On Error GoTo CleanFail

    Dim cht As Chart

    Set cht = GetTargetChart(xlColumnClustered)
    If cht Is Nothing Then Exit Sub

    ApplyChartPipeline cht, "FILL"
    Call RemoveShadow(cht)

    If cht.HasAxis(xlCategory) Then
        cht.Axes(xlCategory).MajorTickMark = xlTickMarkNone
        cht.Axes(xlCategory).MinorTickMark = xlTickMarkNone
    End If

    cht.ChartGroups(1).Overlap = seriesOverlap
    cht.ChartGroups(1).GapWidth = seriesGapWidth
    Exit Sub
CleanFail:
    MsgError "BuildColumnChart"
End Sub


Private Sub BuildStackedColumnChart()
    On Error GoTo CleanFail

    Dim cht As Chart

    Set cht = GetTargetChart(xlColumnStacked)
    If cht Is Nothing Then Exit Sub

    ApplyChartPipeline cht, "FILL"
    Call RemoveShadow(cht)

    If cht.HasAxis(xlCategory) Then
        cht.Axes(xlCategory).MajorTickMark = xlTickMarkNone
        cht.Axes(xlCategory).MinorTickMark = xlTickMarkNone
    End If

    cht.ChartGroups(1).Overlap = 100
    cht.ChartGroups(1).GapWidth = seriesGapWidth
    Exit Sub
CleanFail:
    MsgError "BuildStackedColumnChart"
End Sub


Sub ColumnChart()
    BuildColumnChart
End Sub

Sub StackedColumnChart()
    BuildStackedColumnChart
End Sub
