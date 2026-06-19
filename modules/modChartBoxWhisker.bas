Attribute VB_Name = "modChartBoxWhisker"
'==== Module: modChartBoxWhisker ====
' Box & whisker chart builder (xlBoxwhisker).
'
' Why a separate module
' ---------------------
' Box & whisker is a "chartex" type (Excel 2016+) and uses the worksheet-chrome
' pipeline in modEngineExChrome, not the classic in-chart pipeline: xlBoxwhisker
' rejects cht.Shapes.Add* with error 1004, so its chrome (title/subtitle/figure/
' source/logo + an optional Y-axis title) is built as grouped WORKSHEET shapes.
'
' Unlike treemap, box & whisker DOES have a value axis and a category axis. Native
' Excel tick labels and category labels stay visible inside the chart; only the
' Y-axis *title* is part of the worksheet chrome (defaults.ShowYAxisTitle = True).
' The boxes are real series, so they are coloured per-series from the brand palette
' via FormatSeriesColors (NOT the per-point ApplySliceColors that treemap uses).
'
' Per-type config (Y-axis title on, category strip reserved) is passed to the shared
' chrome via BoxWhiskerChartDefaults / PositionChartExChart - the module does not fork
' the chrome. See modChartTreemap for the reference chartex builder.
Option Explicit


' ============================================================
'   BUILDER
' ============================================================

Private Sub BuildBoxWhiskerChart()
    On Error GoTo CleanFail
    AppFast

    Dim cht As Chart

    ' xlBoxwhisker requires Excel 2016+
    Set cht = GetTargetChart(xlBoxwhisker)
    If cht Is Nothing Then GoTo CleanExit

    Call BuildBoxWhiskerChartWithDefaults(cht, BoxWhiskerChartDefaults())

CleanExit:
    AppRestore
    Exit Sub
CleanFail:
    AppRestore
    MsgError "BuildBoxWhiskerChart"
End Sub

Private Sub BuildBoxWhiskerChartWithDefaults(cht As Chart, ByRef defaults As ChartDefaults)
    On Error GoTo CleanFail

    ' Chrome is built on the host worksheet (see modEngineExChrome) because
    ' xlBoxwhisker rejects cht.Shapes.Add* with error 1004. That requires an embedded
    ' chart with a host worksheet - chart sheets are unsupported.
    If TypeName(cht.Parent) <> "ChartObject" Then
        MsgChartExNeedsEmbedded
        Exit Sub
    End If

    ' Resolve the canvas origin (top-left of the 600x600 layout) BEFORE moving the
    ' chart or rebuilding chrome - on a re-run it reuses the existing canvas position
    ' so the layout stays put even if the chart band has been moved.
    Dim originLeft As Double, originTop As Double
    ChartExCanvasOrigin cht, originLeft, originTop

    ' Position the chart as a band INSIDE the canvas. Box & whisker has a value axis
    ' (showY - leaves room for the worksheet Y-axis title above) AND a category axis
    ' (showX - reserves the in-chart category-label strip below); no legend band is
    ' reserved here. Acts on the ChartObject container, so it is safe and kept OUT of
    ' the defensive block: a positioning failure must surface, not be masked.
    PositionChartExChart cht, originLeft, originTop, showY:=True, showX:=True, hasLegend:=False

    ' Cosmetic chart-object styling ARE classic chart operations that xlBoxwhisker (a
    ' chartex type) can reject with 1004, so apply each defensively and skip rather
    ' than abort - the worksheet chrome is what matters. Make the chart area, plot area
    ' and the ChartObject container transparent with no border so the white canvas
    ' behind shows through cleanly. The chart title is deleted (chrome supplies it);
    ' the legend is LEFT in place (box & whisker has real multi-series legends).
    On Error Resume Next
    cht.ChartArea.Font.name = fontPrimary
    cht.ChartArea.Format.Fill.Visible = msoFalse
    cht.ChartArea.Format.Line.Visible = msoFalse
    cht.ChartArea.Border.LineStyle = xlNone
    cht.PlotArea.Format.Fill.Visible = msoFalse
    cht.PlotArea.Format.Line.Visible = msoFalse
    ' The white box + border is the ChartObject CONTAINER's own ShapeRange, not the cx
    ' chart area: on a chartex chart the ChartArea/PlotArea paths above silently no-op,
    ' so the container fill/line must be cleared too (mirrors the treemap fix).
    With cht.Parent.ShapeRange
        .Fill.Visible = msoFalse
        .Line.Visible = msoFalse
    End With
    If cht.HasTitle Then cht.ChartTitle.Delete   ' chrome supplies the title
    On Error GoTo CleanFail

    ' Colour the boxes from the brand palette. Box & whisker boxes are real series
    ' (one box per category per series), so colour PER-SERIES via FormatSeriesColors -
    ' not the per-point ApplySliceColors that treemap/sunburst use. silent: xlBoxwhisker
    ' may reject standard per-series colouring, and the worksheet chrome is what matters.
    FormatSeriesColors cht, "FILL", silent:=True

    ' Bring data labels onto the house font (Calibri 18pt, brand-3) for every series.
    ' Try series-level DataLabels first; fall back to per-point if the chartex type
    ' rejects it. Guarded so a rejection skips rather than aborts.
    Dim s As Long, seriesCount As Long
    On Error Resume Next
    seriesCount = cht.SeriesCollection.Count
    On Error GoTo CleanFail
    For s = 1 To seriesCount
        On Error Resume Next
        With cht.SeriesCollection(s).DataLabels.Font
            .name = fontPrimary
            .Size = axisFontSize
            .Color = colorBrand3
        End With
        If Err.Number <> 0 Then
            Err.Clear
            Dim p As Long
            For p = 1 To cht.SeriesCollection(s).Points.Count
                With cht.SeriesCollection(s).Points(p).DataLabel.Font
                    .name = fontPrimary
                    .Size = axisFontSize
                    .Color = colorBrand3
                End With
            Next p
        End If
        On Error GoTo CleanFail
    Next s

    ' Title/subtitle/figure/source/logo + white canvas + worksheet Y-axis title as
    ' grouped worksheet shapes, laid out from the same canvas origin.
    ' defaults.ShowYAxisTitle is True for box & whisker (it has a value axis).
    BuildChartExChrome cht, originLeft, originTop, defaults

    Exit Sub
CleanFail:
    MsgError "BuildBoxWhiskerChartWithDefaults"
End Sub


' ============================================================
'   PUBLIC ENTRY POINT
' ============================================================

Sub BoxWhiskerChart()
    BuildBoxWhiskerChart
End Sub
