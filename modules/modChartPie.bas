Attribute VB_Name = "modChartPie"
'==== Module: modChartPie ====
' Pie, donut and treemap chart variants.
'
' Variants
' --------
'   PieChart     - xlPie:      solid filled circle divided into slices
'   DonutChart   - xlDoughnut: ring divided into slices (pie with a hollow centre)
'   TreemapChart - xlTreemap:  hierarchical rectangular tiles
'
' Differences
' -----------
'   Pie/Donut: custom pipeline using SetRoundChartSizeAndTitle; slice colours from brand
'              palette. Pie and donut share the same builder (BuildPieChartWithDefaults) -
'              the round-chart sizing and slice colouring are identical; only the chart
'              type differs (xlPie vs xlDoughnut).
'   Treemap: custom pipeline; no axes or gridlines. Tiles are points of a single
'              series, so they are coloured per-point from the brand palette via
'              the same ApplySliceColors helper that pie/donut use. xlTreemap rejects
'              cht.Shapes.Add*, so its chrome (title/subtitle/logo/source) is built as
'              grouped WORKSHEET shapes by modTreemapChrome - not inside the chart.
'
' Pie/Donut use a custom pipeline (no ApplyChartPipeline) because they have no axes
' or gridlines. Steps applied: InsertSource, SetRoundChartSizeAndTitle (which calls
' FormatTitle), InsertLogo, slice colouring.
'
' Palette: 8 data colours (Teal, Jasmine, Navy, Coral, Sky, Cherry, Blush, Indigo).
' Slices beyond 8 use colorNeutral2 (Steel).
Option Explicit


' ============================================================
'   BUILDERS
' ============================================================

Private Sub BuildPieChart()
    On Error GoTo CleanFail
    AppFast

    Dim cht As Chart

    Set cht = GetTargetChart(xlPie)
    If cht Is Nothing Then GoTo CleanExit

    Call BuildPieChartWithDefaults(cht, PieChartDefaults())

CleanExit:
    AppRestore
    Exit Sub
CleanFail:
    AppRestore
    MsgError "BuildPieChart"
End Sub

Private Sub BuildDonutChart()
    On Error GoTo CleanFail
    AppFast

    Dim cht As Chart

    ' Donut shares the pie builder - only the chart type differs (xlDoughnut vs xlPie).
    Set cht = GetTargetChart(xlDoughnut)
    If cht Is Nothing Then GoTo CleanExit

    Call BuildPieChartWithDefaults(cht, PieChartDefaults())

CleanExit:
    AppRestore
    Exit Sub
CleanFail:
    AppRestore
    MsgError "BuildDonutChart"
End Sub

Private Sub BuildPieChartWithDefaults(cht As Chart, ByRef defaults As ChartDefaults)
    On Error GoTo CleanFail

    Dim pointscount As Long

    InsertSource cht
    SetRoundChartSizeAndTitle cht, defaults
    InsertLogo cht      ' must follow SetRoundChartSizeAndTitle so chart is 600×600 when logo is sized

    pointscount = cht.SeriesCollection(1).Points.Count
    ApplySliceColors cht, pointscount

    Exit Sub
CleanFail:
    MsgError "BuildPieChartWithDefaults"
End Sub


' ============================================================
'   SHARED PRIVATE HELPERS
' ============================================================

' Public so the chart-type-agnostic styler (modChartStyle.ApplyChartStyle) can
' reuse per-slice colouring instead of duplicating the loop.
' silent: see InsertLogo - suppresses the failure message for the agnostic
' styler, where a chart type (e.g. treemap) may reject per-point colouring.
Public Sub ApplySliceColors(cht As Chart, ByVal pointscount As Long, Optional ByVal silent As Boolean = False)
    On Error GoTo CleanFail

    Dim i As Long
    Dim sliceColor As Long

    For i = 1 To pointscount
        ' GetPaletteColor returns the brand palette color for i <= 8.
        ' For slices beyond 8, use colorNeutral2 (Steel) instead of the default colorNeutral1.
        If i <= 8 Then
            sliceColor = GetPaletteColor(i)
        Else
            sliceColor = colorNeutral2
        End If

        With cht.SeriesCollection(1).Points(i).Format
            With .Fill
                .Visible = msoTrue
                .Solid
                .ForeColor.RGB = sliceColor
            End With
            .Line.Visible = msoFalse
        End With
    Next i

    Exit Sub
CleanFail:
    If Not silent Then MsgError "ApplySliceColors"
End Sub


' Public so the chart-type-agnostic styler (modChartStyle.ApplyChartStyle) can
' reuse pie/donut square-plot sizing + title instead of the rectangular pipeline.
Public Sub SetRoundChartSizeAndTitle(cht As Chart, ByRef defaults As ChartDefaults)
    On Error GoTo CleanFail

    ' Shared layout for both pie and donut - chart dimensions, text boxes, plot area
    ' sizing, centering, and legend placement are identical for both variants.
    Dim chtObj As ChartObject
    Dim chtHeight As Double
    Dim chtWidth As Double
    Dim pltWidth As Double
    Dim pltHeight As Double
    Dim plotSize As Long

    With cht.Parent
        .Width = chartWidth
        .Height = chartHeight
    End With

    cht.ChartArea.Font.name = fontPrimary
    cht.ChartArea.Border.LineStyle = xlNone

    FormatTitle cht

    plotSize = IIf(cht.hasLegend, pieplotAreaSize_legend, pieplotAreaSize_noLegend)
    With cht.PlotArea
        .Width = plotSize
        .Height = plotSize
        .Left = pieplotAreaLeft
        .Top = pieplotAreaTop
    End With

    Set chtObj = cht.Parent
    With chtObj
        chtHeight = .Chart.ChartArea.Height
        chtWidth = .Chart.ChartArea.Width
        pltHeight = .Chart.PlotArea.Height
        pltWidth = .Chart.PlotArea.Width
        .Chart.PlotArea.Top = (chtHeight - pltHeight) * piePlotTopRatio
        .Chart.PlotArea.Left = (chtWidth - pltWidth) / 2
    End With

    If cht.hasLegend Then
        With cht.Legend
            .Position = xlLegendPositionTop
            .Left = legendLeftPad
            .Font.Color = legendFontColor
            .Top = pieLegendTop
            .Font.Size = axisFontSize
        End With
    End If

    Exit Sub
CleanFail:
    MsgError "SetRoundChartSizeAndTitle"
End Sub


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

    ' Treemap chrome is built on the host worksheet (see modTreemapChrome) because
    ' xlTreemap rejects cht.Shapes.Add* with error 1004. That requires an embedded
    ' chart with a host worksheet - chart sheets are unsupported.
    If TypeName(cht.Parent) <> "ChartObject" Then
        MsgTreemapNeedsEmbedded
        Exit Sub
    End If

    ' Resolve the canvas origin (top-left of the 600x600 layout) BEFORE moving the
    ' chart or rebuilding chrome - on a re-run it reuses the existing canvas position
    ' so the layout stays put even if the chart band has been moved.
    Dim originLeft As Double, originTop As Double
    TreemapCanvasOrigin cht, originLeft, originTop

    ' Position the chart as a band INSIDE the canvas, leaving room for the title block
    ' above and the logo/source below (mirrors the classic plot-area geometry). This
    ' acts on the ChartObject container (not the chartex chart), so it is safe and is
    ' kept OUT of the defensive block: a positioning failure must surface, not be
    ' masked into a misplaced layout over a full-size chart.
    PositionTreemapChart cht, originLeft, originTop

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

    ' Title/subtitle/figure/source/logo + white canvas as grouped worksheet shapes,
    ' laid out from the same canvas origin.
    BuildTreemapChrome cht, originLeft, originTop

    Exit Sub
CleanFail:
    MsgError "BuildTreemapChartWithDefaults"
End Sub


' ============================================================
'   PUBLIC ENTRY POINTS
' ============================================================

Sub PieChart()
    BuildPieChart
End Sub

Sub DonutChart()
    BuildDonutChart
End Sub

Sub TreemapChart()
    BuildTreemapChart
End Sub
