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
' Palette: first 5 brand colours (Ocean, Coral, Sky, Pine, Gold).
' Maximum 5 slices; charts with more will show a warning and receive no colouring.
Option Explicit


' ============================================================
'   BUILDERS
' ============================================================

Private Sub BuildPieChart()
    Dim cht As Chart
    Dim pointscount As Long

    Set cht = GetTargetChart(xlPie)
    If cht Is Nothing Then Exit Sub

    InsertSource cht
    SetRoundChartSizeAndTitle cht
    InsertLogo cht      ' must follow SetRoundChartSizeAndTitle so chart is 600×600 when logo is sized

    pointscount = cht.SeriesCollection(1).Points.Count
    ApplySliceColors cht, pointscount
End Sub


' ============================================================
'   SHARED PRIVATE HELPERS
' ============================================================

Private Sub ApplySliceColors(cht As Chart, ByVal pointscount As Long)
    If pointscount > 5 Then
        MsgTooManySeries cht
        Exit Sub
    End If

    Dim palette(1 To 5) As Long
    palette(1) = colorData1
    palette(2) = colorData2
    palette(3) = colorData3
    palette(4) = colorData4
    palette(5) = colorData5

    Dim i As Long
    For i = 1 To pointscount
        With cht.SeriesCollection(1).Points(i).Format
            With .Fill
                .Visible = msoTrue
                .Solid
                .ForeColor.rgb = palette(i)
            End With
            .Line.Visible = msoFalse
        End With
    Next i
End Sub


Private Sub SetRoundChartSizeAndTitle(cht As Chart)
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

    cht.ChartArea.Font.Name = fontPrimary
    cht.ChartArea.Border.LineStyle = xlNone

    FormatTitle cht

    plotSize = IIf(cht.HasLegend, piePlotAreaSize_legend, piePlotAreaSize_noLegend)
    cht.PlotArea.Select
    Selection.Width = plotSize
    Selection.Height = plotSize
    Selection.Left = piePlotAreaLeft_web
    Selection.Top = piePlotAreaTop_web

    Set chtObj = cht.Parent
    With chtObj
        chtHeight = .Chart.ChartArea.Height
        chtWidth = .Chart.ChartArea.Width
        pltHeight = .Chart.PlotArea.Height
        pltWidth = .Chart.PlotArea.Width
        .Chart.PlotArea.Top = (chtHeight - pltHeight) * piePlotTopRatio_web
        .Chart.PlotArea.Left = (chtWidth - pltWidth) / 2
    End With

    If cht.HasLegend Then
        cht.Legend.Position = xlLegendPositionTop
        cht.Legend.Select
        Selection.Top = pieLegendTop_web
        Selection.Font.Size = axisFontSize
    End If
End Sub


Private Sub BuildTreemapChart()
    Dim cht As Chart

    ' xlTreemap requires Excel 2016+
    Set cht = GetTargetChart(xlTreemap)
    If cht Is Nothing Then Exit Sub

    ' Custom pipeline — treemaps have no axes or gridlines
    With cht.Parent
        .Width = chartWidth
        .Height = chartHeight
    End With
    cht.ChartArea.Font.Name = fontPrimary
    cht.ChartArea.Border.LineStyle = xlNone

    ' Tile labels make a legend redundant
    If cht.HasLegend Then cht.Legend.Delete

    InsertSource cht
    FormatTitle cht
    InsertLogo cht  ' must follow size-setting so logo is sized against 600x600
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
