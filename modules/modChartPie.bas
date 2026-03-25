Attribute VB_Name = "modChartPie"
'==== Module: modChartPie ====
' Pie and treemap chart variants.
'
' Variants
' --------
'   PieChart     — xlPie:     solid filled circle divided into slices
'   TreemapChart — xlTreemap: hierarchical rectangular tiles
'
' Differences
' -----------
'   Pie: custom pipeline using SetRoundChartSizeAndTitle; slice colours from brand palette.
'   Treemap: custom pipeline; no axes or gridlines; tile colours managed by Excel.
'
' Pie uses a custom pipeline (no ApplyChartPipeline) because pie charts have no axes
' or gridlines. Steps applied: InsertSource, SetRoundChartSizeAndTitle (which calls
' FormatTitle), InsertLogo, slice colouring.
'
' To toggle between pie and donut, use ToggleChartVariant (modChartTools).
'
' Palette: 7 data colours (Ocean, Coral, Sky, Pine, Gold, Rust, Lavender).
' Slices beyond 7 use colorNeutral2 (Steel).
Option Explicit


' ============================================================
'   BUILDERS
' ============================================================

Private Sub BuildPieChart()
    On Error GoTo CleanFail

    Dim cht As Chart

    Set cht = GetTargetChart(xlPie)
    If cht Is Nothing Then Exit Sub

    Call BuildPieChartWithDefaults(cht, PieChartDefaults())

    Exit Sub
CleanFail:
    MsgError "BuildPieChart"
End Sub

Private Sub BuildPieChartWithDefaults(cht As Chart, ByVal defaults As ChartDefaults)
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

Private Sub ApplySliceColors(cht As Chart, ByVal pointscount As Long)
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
    MsgError "ApplySliceColors"
End Sub


Private Sub SetRoundChartSizeAndTitle(cht As Chart, ByVal defaults As ChartDefaults)
    On Error GoTo CleanFail

    ' Shared layout for both pie and donut — chart dimensions, text boxes, plot area
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

    plotSize = IIf(cht.hasLegend, piePlotAreaSize_legend, piePlotAreaSize_noLegend)
    cht.PlotArea.Select
    Selection.Width = plotSize
    Selection.Height = plotSize
    Selection.Left = piePlotAreaLeft
    Selection.Top = piePlotAreaTop

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
        cht.Legend.Position = xlLegendPositionTop
        cht.Legend.Left = legendLeftPad
        cht.Legend.Font.Color = legendFontColor
        cht.Legend.Select
        Selection.Top = pieLegendTop
        Selection.Font.Size = axisFontSize
    End If

    Exit Sub
CleanFail:
    MsgError "SetRoundChartSizeAndTitle"
End Sub


Private Sub BuildTreemapChart()
    On Error GoTo CleanFail

    Dim cht As Chart

    ' xlTreemap requires Excel 2016+
    Set cht = GetTargetChart(xlTreemap)
    If cht Is Nothing Then Exit Sub

    Call BuildTreemapChartWithDefaults(cht, TreemapChartDefaults())

    Exit Sub
CleanFail:
    MsgError "BuildTreemapChart"
End Sub

Private Sub BuildTreemapChartWithDefaults(cht As Chart, ByVal defaults As ChartDefaults)
    On Error GoTo CleanFail

    ' Custom pipeline — treemaps have no axes or gridlines
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

Sub TreemapChart()
    BuildTreemapChart
End Sub
