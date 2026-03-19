Attribute VB_Name = "modChartScatter"
Option Explicit

Private Sub BuildScatterChart()
    Dim cht As Chart
    Dim seriescount As Long
    Dim imarker As Long

    ActiveSheet.Shapes.AddChart2(-1, xlXYScatter).Select
    ActiveChart.Parent.Duplicate.Select

    Set cht = ActiveChart
    If cht Is Nothing Then Exit Sub

    ' Scatter uses a partial pipeline (ScatterplotStyles handles gridlines)
    OuterFormat cht
    FormatXAxisTitle cht
    InsertLogo cht
    InsertSource cht
    FormatTitle cht
    FormatXAxis cht
    ScatterplotStyles cht

    ' Scatter-specific: tick marks outside on category axis
    cht.Axes(xlCategory).MajorTickMark = xlTickMarkOutside
    cht.Axes(xlCategory).MinorTickMark = xlTickMarkNone

    seriescount = cht.SeriesCollection.Count

    ' Set marker style and size for all series
    For imarker = 1 To seriescount
        cht.SeriesCollection(imarker).Select
        With Selection
            .MarkerStyle = xlMarkerStyleCircle
            .MarkerSize = scatterMarkerSize
        End With
    Next imarker

    ' Apply brand palette to markers (up to 7 series)
    If seriescount > 7 Then
        MsgTooManySeries cht
    Else
        Dim palette(1 To 7) As Long
        palette(1) = colorOcean
        palette(2) = colorCoral
        palette(3) = colorSky
        palette(4) = colorPine
        palette(5) = colorGold
        palette(6) = colorRust
        palette(7) = colorLavender

        Dim i As Long
        For i = 1 To seriescount
            With cht.SeriesCollection(i)
                .Border.ColorIndex = xlNone
                .MarkerForegroundColorIndex = xlColorIndexNone
                .MarkerBackgroundColor = palette(i)
            End With
        Next i
    End If
End Sub


Private Sub ScatterplotStyles(cht As Chart)
    ' Format vertical gridlines
    If Not cht.Axes(xlCategory).HasMajorGridlines Then
        cht.SetElement (msoElementPrimaryCategoryGridLinesMajor)
    End If
    cht.Axes(xlCategory).MajorGridlines.Select
    With Selection.Format.Line
        .Visible = msoTrue
        .Weight = gridlineWeight
        .DashStyle = msoLineSysDot
        .ForeColor.rgb = giRGBgridlinesweb
    End With

    ' Format horizontal gridlines
    If Not cht.Axes(xlValue).HasMajorGridlines Then
        cht.PlotArea.Select
        cht.SetElement (msoElementPrimaryValueGridLinesMajor)
    End If
    cht.Axes(xlValue).MajorGridlines.Select
    With Selection.Format.Line
        .Visible = msoTrue
        .Weight = gridlineWeight
        .DashStyle = msoLineSysDot
        .ForeColor.rgb = giRGBgridlinesweb
    End With
End Sub


Sub ScatterChart()
    BuildScatterChart
End Sub

Public Sub Scatter_onAction(control As IRibbonControl)
    BuildScatterChart
End Sub
