Attribute VB_Name = "modChartPie"
Option Explicit

Private Sub BuildPieChart()
    Dim cht As Chart
    Dim pointscount As Long

    ActiveSheet.Shapes.AddChart2(-1, xlPie).Select
    ActiveChart.Parent.Duplicate.Select

    Set cht = ActiveChart
    If cht Is Nothing Then Exit Sub

    ' Pie-specific pipeline (no generic ApplyChartPipeline — pie has no axes/gridlines)
    InsertLogo cht
    InsertSource cht
    SetPieChartSizeandTitle cht

    ' Remove chart border
    cht.ChartArea.Border.LineStyle = xlNone

    ' Color pie slices
    pointscount = cht.SeriesCollection(1).Points.Count

    If pointscount > 5 Then
        MsgTooManySeries cht
    Else
        Dim palette(1 To 5) As Long
        palette(1) = colorOcean
        palette(2) = colorCoral
        palette(3) = colorSky
        palette(4) = colorPine
        palette(5) = colorGold

        Dim i As Long
        For i = 1 To pointscount
            ApplyPieSliceColor cht, i, palette(i)
        Next i
    End If
End Sub


Private Sub ApplyPieSliceColor(cht As Chart, ByVal idx As Long, ByVal clr As Long)
    With cht.SeriesCollection(1).Points(idx).Format
        With .Fill
            .Visible = msoTrue
            .Solid
            .ForeColor.rgb = clr
        End With
        .Line.Visible = msoFalse
    End With
End Sub


Private Sub SetPieChartSizeandTitle(cht As Chart)
    Dim titleB1 As TextBox
    Dim titleB2 As TextBox
    Dim titleB3 As TextBox
    Dim chtObj As ChartObject
    Dim chtHeight As Double
    Dim chtWidth As Double
    Dim pltWidth As Double
    Dim pltHeight As Double
    Dim plotSize As Long

    ' Set chart dimensions
    With cht.Parent
        .Width = chartWidth
        .Height = chartHeight
    End With

    cht.ChartArea.Font.Name = fontPrimary

    ' Remove built-in title; replace with text boxes
    If cht.HasTitle Then cht.ChartTitle.Delete

    Set titleB1 = cht.TextBoxes.Add(0, 0, titleBoxWidth, pieTitleBoxHeight)
    With titleB1
        .Name = "TitleBox"
        .Text = "Title in 18pt Title Case"
        .Font.Size = pieTitleFontSize
        .Font.Name = fontPrimary
        .Font.Bold = msoFalse
    End With

    Set titleB2 = cht.TextBoxes.Add(0, pieSubtitleBoxTop, titleBoxWidth, pieSubtitleBoxHeight)
    With titleB2
        .Name = "SubTitleBox"
        .Text = "Subtitle in 14pt sentence case"
        .Font.Size = pieSubtitleFontSize
        .Font.Name = fontPrimary
        .Font.Bold = msoFalse
    End With

    Set titleB3 = cht.TextBoxes.Add(0, pieYAxisLabelBoxTop, titleBoxWidth, yAxisLabel_noLegendHeight)
    With titleB3
        .Name = "YAxisLabelBox"
        .Text = "Y axis title (unit)"
        .Font.Size = axisFontSize
        .Font.Name = fontPrimaryItalic
        .Font.Bold = msoFalse
        .Font.Italic = msoTrue
    End With

    ' Size the pie chart plot area
    plotSize = IIf(cht.HasLegend, piePlotAreaSize_legend, piePlotAreaSize_noLegend)
    cht.PlotArea.Select
    Selection.Width = plotSize
    Selection.Height = plotSize
    Selection.Left = piePlotAreaLeft_web
    Selection.Top = piePlotAreaTop_web

    ' Center pie chart horizontally
    Set chtObj = cht.Parent
    With chtObj
        chtHeight = .Chart.ChartArea.Height
        chtWidth = .Chart.ChartArea.Width
        pltHeight = .Chart.PlotArea.Height
        pltWidth = .Chart.PlotArea.Width
        .Chart.PlotArea.Top = (chtHeight - pltHeight) * piePlotTopRatio_web
        .Chart.PlotArea.Left = (chtWidth - pltWidth) / 2
    End With

    ' Position the legend
    If cht.HasLegend Then
        cht.Legend.Position = xlLegendPositionTop
        cht.Legend.Select
        Selection.Top = pieLegendTop_web
        Selection.Font.Size = axisFontSize
    End If
End Sub


Sub PieChart()
    BuildPieChart
End Sub
