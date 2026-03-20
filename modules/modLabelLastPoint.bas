Attribute VB_Name = "modLabelLastPoint"
Option Explicit

Private Sub BuildLabelLastPoint()
    Dim ipts As Long
    Dim Npts As Long
    Dim bLabeled As Boolean
    Dim cht As Chart
    Dim srs As Series
    Dim plHeight As Double
    Dim plWidth As Double
    Dim shp As Shape
    Dim iColor As Long
    Dim lbl As DataLabel

    If ActiveChart Is Nothing Then
        MsgNoActiveChart
        Exit Sub
    End If

    ' Duplicate and activate the copy
    ActiveChart.Parent.Duplicate.Select
    ActiveChart.PlotArea.Select

    ' Narrow plot area only on line charts to make room for end labels
    If ActiveChart.ChartType = xlLine Then
        Selection.Width = chartWidth - labelLastPointPlotWidthInset
    End If
    Selection.Left = 0

    Set cht = ActiveChart

    ' Nudge Y-axis label box upward when legend is present
    If cht.HasLegend Then
        For Each shp In cht.Shapes
            If shp.Name = "YAxisLabelBox" Then
                cht.Shapes.Range(Array("YAxisLabelBox")).Select
                Selection.ShapeRange.IncrementTop labelLastPointTitleNudge
            End If
        Next shp
    End If

    ' Adjust plot area dimensions
    cht.PlotArea.Select
    plHeight = cht.PlotArea.Height
    plWidth = cht.PlotArea.Width
    Selection.Top = labelLastPointPlotTop
    Selection.Width = plWidth * labelLastPointPlotWidthRatio_web
    Selection.Height = plHeight

    ' Remove legend (labels replace it)
    If cht.HasLegend Then
        cht.Legend.Delete
    End If

    ' Label the last valid point in each series
    For Each srs In cht.SeriesCollection
        bLabeled = False
        With srs
            Npts = 0
            On Error Resume Next
            Npts = .Points.Count
            On Error GoTo 0

            If Npts > 0 Then
                For ipts = Npts To 1 Step -1
                    On Error Resume Next
                    If bLabeled Then
                        srs.Points(ipts).HasDataLabel = False
                    Else
                        ' Clear any existing label first (linked labels resist reassignment)
                        srs.Points(ipts).HasDataLabel = False
                        srs.Points(ipts).ApplyDataLabels _
                            ShowSeriesName:=True, ShowCategoryName:=False, _
                            ShowValue:=False, AutoText:=False, LegendKey:=False
                        bLabeled = (Err.Number = 0)
                        ' Excel 2010+: no error on unplotted points but label is blank
                        If bLabeled Then bLabeled = (Len(srs.Points(ipts).DataLabel.Text) > 0)
                        If Not bLabeled Then srs.Points(ipts).HasDataLabel = False
                    End If
                    On Error GoTo 0

                    If bLabeled Then
                        Set lbl = srs.Points(ipts).DataLabel
                        lbl.Font.Bold = msoTrue

                        Select Case srs.ChartType
                            Case xlLine, xlLineStacked, xlLineStacked100, xlLineMarkers, xlLineMarkersStacked, xlLineMarkersStacked100
                                lbl.Position = xlLabelPositionRight
                                iColor = .Format.Line.ForeColor.rgb
                            Case xlXYScatter, xlXYScatterLines, xlXYScatterLinesNoMarkers, xlXYScatterSmooth, xlXYScatterSmoothNoMarkers
                                lbl.Position = xlLabelPositionRight
                                iColor = .MarkerBackgroundColor
                            Case xlColumnClustered, xlBarClustered
                                lbl.Position = xlLabelPositionOutsideEnd
                                iColor = .Format.Fill.ForeColor.rgb
                            Case xlColumnStacked, xlColumnStacked100, xlBarStacked, xlBarStacked100, xlArea, xlAreaStacked, xlAreaStacked100
                                lbl.Position = xlLabelPositionCenter
                                iColor = .Format.Fill.ForeColor.rgb
                        End Select

                        lbl.Font.Color = iColor
                        lbl.Font.Size = axisFontSize
                    End If
                Next ipts
            End If

            ' Required so label updates when series name changes
            srs.DataLabels.AutoText = True
        End With
    Next srs
End Sub


Sub LabelLastPointButton()
    BuildLabelLastPoint
End Sub
