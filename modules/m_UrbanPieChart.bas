Attribute VB_Name = "m_UrbanPieChart"
Option Explicit

Private Sub BuildPieChart()
    Dim cht As Chart
    Dim txtB As Shape
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
        Set txtB = cht.Shapes.AddTextbox(msoTextOrientationHorizontal, 0, 0, pieErrorBoxWidth, pieErrorBoxHeight)
        With txtB
            .Name = "TitleBox"
            With .TextFrame2.TextRange
                .Text = "You have too many data series for this chart type. Please contact the Communications Department for further guidance."
                .Font.Size = pieErrorFontSize
                .Font.Name = FontStyle
                .Font.Fill.ForeColor.rgb = vbRed
                .ParagraphFormat.Alignment = msoTextEffectAlignmentLeft
            End With
            .Fill.ForeColor.rgb = vbYellow
        End With
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


Sub UrbanPieChart()
    BuildPieChart
End Sub

Public Sub Pie_onAction(control As IRibbonControl)
    BuildPieChart
End Sub
