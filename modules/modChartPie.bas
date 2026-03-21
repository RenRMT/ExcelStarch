Attribute VB_Name = "modChartPie"
'==== Module: modChartPie ====
' Pie and donut chart variants.
'
' Variants
' --------
'   PieChart   — xlPie:      solid filled circle divided into slices
'   DonutChart — xlDoughnut: same as pie with a hollow centre
'
' Differences
' -----------
'   Chart type only: xlPie vs xlDoughnut.
'   All sizing, layout, title boxes, legend handling, slice colouring, and
'   border removal are identical and share the private helpers below.
'
' Both variants use a custom pipeline (no ApplyChartPipeline) because pie/donut
' charts have no axes or gridlines. Steps applied: InsertLogo, InsertSource,
' SetRoundChartSizeAndTitle (which calls FormatTitle), slice colouring.
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

    InsertLogo cht
    InsertSource cht
    SetRoundChartSizeAndTitle cht

    pointscount = cht.SeriesCollection(1).Points.Count
    ApplySliceColors cht, pointscount
End Sub


Private Sub BuildDonutChart()
    Dim cht As Chart
    Dim pointscount As Long

    ' xlDoughnut — identical to pie in all respects except the hollow centre.
    ' Hole size is controlled by Excel's default (75%); no VBA override applied.
    Set cht = GetTargetChart(xlDoughnut)
    If cht Is Nothing Then Exit Sub

    InsertLogo cht
    InsertSource cht
    SetRoundChartSizeAndTitle cht

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


' ============================================================
'   PUBLIC ENTRY POINTS
' ============================================================

Sub PieChart()
    BuildPieChart
End Sub

Sub DonutChart()
    BuildDonutChart
End Sub
