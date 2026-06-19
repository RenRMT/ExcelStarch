Attribute VB_Name = "modChartTreemap"
'==== Module: modChartTreemap ====
' Treemap chart builder (xlTreemap - hierarchical rectangular tiles).
'
' Why a separate module
' ---------------------
' Treemap is a "chartex" type (Excel 2016+), a different object family from the
' classic charts in modEngineBuilder. It does NOT use ApplyChartPipeline: xlTreemap
' rejects cht.Shapes.Add* with error 1004, so its chrome (title/subtitle/figure/
' source/logo) cannot live inside the chart and is built as grouped WORKSHEET shapes
' by modEngineExChrome instead. It has no axes or gridlines.
'
' Tiles are points of a single series (like pie slices), so they are coloured
' per-point from the brand palette via the shared ApplySliceColors helper (Public in
' modChartPie - reused here cross-module, exactly as modEngineStyle does).
'
' This module owns only the treemap BUILD pipeline; the worksheet chrome, grouping
' and export live in modEngineExChrome and modExport.
Option Explicit


' ============================================================
'   BUILDER
' ============================================================

Private Sub BuildTreemapChart()
    On Error GoTo CleanFail
    AppFast

    Dim cht As Chart

    ' xlTreemap requires Excel 2016+
    Set cht = GetTargetChart(xlTreemap)
    If cht Is Nothing Then GoTo CleanExit

    Call BuildTreemapChartWithDefaults(cht, TreemapChartDefaults())

CleanExit:
    AppRestore
    Exit Sub
CleanFail:
    AppRestore
    MsgError "BuildTreemapChart"
End Sub

Private Sub BuildTreemapChartWithDefaults(cht As Chart, ByRef defaults As ChartDefaults)
    On Error GoTo CleanFail

    Dim pointscount As Long

    ' Treemap chrome is built on the host worksheet (see modEngineExChrome) because
    ' xlTreemap rejects cht.Shapes.Add* with error 1004. That requires an embedded
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

    ' Position the chart as a band INSIDE the canvas, leaving room for the title block
    ' above and the logo/source below (mirrors the classic plot-area geometry). Treemap
    ' has a value-axis-title band above (showY) but no category strip below (showX) and
    ' no legend. This acts on the ChartObject container (not the chartex chart), so it
    ' is safe and is kept OUT of the defensive block: a positioning failure must
    ' surface, not be masked into a misplaced layout over a full-size chart.
    PositionChartExChart cht, originLeft, originTop, showY:=True, showX:=False, hasLegend:=False

    ' Cosmetic chart-object styling and title/legend removal ARE classic chart
    ' operations that xlTreemap (a chartex type) can reject with 1004, so apply each
    ' defensively and skip rather than abort - the worksheet chrome is what matters.
    ' Make the chart area and plot area transparent with no border so the white canvas
    ' behind shows through cleanly (the canvas supplies the background, not the chart).
    On Error Resume Next
    cht.ChartArea.Font.name = fontPrimary
    cht.ChartArea.Format.Fill.Visible = msoFalse
    cht.ChartArea.Format.Line.Visible = msoFalse
    cht.ChartArea.Border.LineStyle = xlNone
    cht.PlotArea.Format.Fill.Visible = msoFalse
    cht.PlotArea.Format.Line.Visible = msoFalse
    ' The white box + border the user sees is the ChartObject CONTAINER's own
    ' ShapeRange, not the cx chart area: on a chartex chart the ChartArea/PlotArea
    ' paths above silently no-op, so the container fill/line must be cleared too.
    ' (If a build refuses transparency here, swap to a white solid fill + no line so
    '  the container blends into the white Canvas behind it instead.)
    With cht.Parent.ShapeRange
        .Fill.Visible = msoFalse
        .Line.Visible = msoFalse
    End With
    If cht.HasTitle Then cht.ChartTitle.Delete   ' chrome supplies the title; chart has none
    If cht.hasLegend Then cht.Legend.Delete       ' tile labels make a legend redundant
    On Error GoTo CleanFail

    ' Colour tiles from the brand palette. Treemap tiles are points of a single
    ' series (like pie slices), so the per-point ApplySliceColors loop applies.
    ' Excel's acceptance of per-point .Fill on xlTreemap is not guaranteed, so
    ' read the point count defensively and skip colouring rather than raising a
    ' misleading error if the type rejects it.
    On Error Resume Next
    pointscount = cht.SeriesCollection(1).Points.Count
    On Error GoTo CleanFail
    If pointscount > 0 Then ApplySliceColors cht, pointscount, silent:=True

    ' Bring tile data labels onto the house font (Calibri 18pt, brand-3) - they are
    ' not styled by ApplySliceColors. Try the whole series in one call; if the chartex
    ' type rejects series-level DataLabels, fall back to styling each tile's label.
    If pointscount > 0 Then
        On Error Resume Next
        With cht.SeriesCollection(1).DataLabels.Font
            .name = fontPrimary
            .Size = axisFontSize
            .Color = colorBrand3
        End With
        If Err.Number <> 0 Then
            Err.Clear
            Dim p As Long
            For p = 1 To pointscount
                With cht.SeriesCollection(1).Points(p).DataLabel.Font
                    .name = fontPrimary
                    .Size = axisFontSize
                    .Color = colorBrand3
                End With
            Next p
        End If
        On Error GoTo CleanFail
    End If

    ' Title/subtitle/figure/source/logo + white canvas as grouped worksheet shapes,
    ' laid out from the same canvas origin. defaults.ShowYAxisTitle is False for treemap
    ' (no value axis), so no Y-axis title box is added.
    BuildChartExChrome cht, originLeft, originTop, defaults

    Exit Sub
CleanFail:
    MsgError "BuildTreemapChartWithDefaults"
End Sub


' ============================================================
'   PUBLIC ENTRY POINT
' ============================================================

Sub TreemapChart()
    BuildTreemapChart
End Sub
