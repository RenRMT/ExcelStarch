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
'              the same ApplySliceColors helper that pie/donut use.
'
' Pie/Donut use a custom pipeline (no ApplyChartPipeline) because they have no axes
' or gridlines. Steps applied: InsertSource, SetRoundChartSizeAndTitle (which calls
' FormatTitle), InsertLogo, slice colouring.
'
' Palette: 7 data colours (Ocean, Coral, Sky, Pine, Gold, Rust, Lavender).
' Slices beyond 7 use colorNeutral2 (Steel).
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
        ' GetPaletteColor returns the brand palette color for i <= 7.
        ' For slices beyond 7, use colorNeutral2 (Steel) instead of the default colorNeutral1.
        If i <= 7 Then
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

    ' Custom pipeline - treemaps have no axes or gridlines
    With cht.Parent
        .Width = chartWidth
        .Height = chartHeight
    End With
    cht.ChartArea.Font.name = fontPrimary
    cht.ChartArea.Border.LineStyle = xlNone

    ' Tile labels make a legend redundant
    If cht.hasLegend Then cht.Legend.Delete

    InsertSource cht
    FormatTitle cht
    InsertLogo cht  ' must follow size-setting so logo is sized against 600x600

    ' Colour tiles from the brand palette. Treemap tiles are points of a single
    ' series (like pie slices), so the per-point ApplySliceColors loop applies.
    ' Excel's acceptance of per-point .Fill on xlTreemap is not guaranteed, so
    ' read the point count defensively and skip colouring rather than raising a
    ' misleading error if the type rejects it.
    On Error Resume Next
    pointscount = cht.SeriesCollection(1).Points.Count
    On Error GoTo CleanFail
    If pointscount > 0 Then ApplySliceColors cht, pointscount

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
