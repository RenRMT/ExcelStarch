Attribute VB_Name = "modChartTools"
'==== Module: modChartTools ====
' Post-creation chart utilities triggered from the Customisation and FillActions
' ribbon groups. All tools operate on an already-formatted chart.
'
' Tools
' -----
'   LabelLastPointButton  — adds series-name labels to the final data point of each
'                           series; duplicates the chart first
'   ToggleGridlines       — cycles major gridlines: None → Horizontal → Vertical → Both
'   RemoveLegendResizeButton — deletes the legend and resizes the plot area to standard
'                           web dimensions; operates in-place
'   StartWithGray         — duplicates the chart and resets all series to Silver;
'                           GrayOutChart is the parameterised core (also public for
'                           potential pipeline use)
'
' Duplication behaviour
' ---------------------
'   LabelLastPoint and StartWithGray both duplicate the source chart by default so the
'   original is preserved. ToggleGridlines and RemoveLegendResize operate in-place on
'   the active chart — they are intended for iterative adjustment, not one-shot creation.
Option Explicit


' ============================================================
'   LABEL LAST POINT
' ============================================================

Private Sub BuildLabelLastPoint()
    On Error GoTo CleanFail

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
    Exit Sub

CleanFail:
    MsgError "BuildLabelLastPoint"
End Sub

Sub LabelLastPointButton()
    BuildLabelLastPoint
End Sub


' ============================================================
'   TOGGLE GRIDLINES
' ============================================================
' Cycles major gridlines through four states in sequence:
'   None → Horizontal only → Vertical only → Both → None
' Operates in-place on the active chart (no duplication).

Public Sub ToggleGridlines()
    If ActiveChart Is Nothing Then
        MsgNoActiveChart
        Exit Sub
    End If

    Dim cht As Chart
    Set cht = ActiveChart

    Dim hasH As Boolean     ' horizontal gridlines (value / Y axis)
    Dim hasV As Boolean     ' vertical gridlines   (category / X axis)

    If cht.HasAxis(xlValue) Then hasH = cht.Axes(xlValue).HasMajorGridlines
    If cht.HasAxis(xlCategory) Then hasV = cht.Axes(xlCategory).HasMajorGridlines

    Dim nextH As Boolean
    Dim nextV As Boolean

    If Not hasH And Not hasV Then
        nextH = True:  nextV = False        ' None → Horizontal only
    ElseIf hasH And Not hasV Then
        nextH = False: nextV = True         ' Horizontal → Vertical only
    ElseIf Not hasH And hasV Then
        nextH = True:  nextV = True         ' Vertical → Both
    Else
        nextH = False: nextV = False        ' Both → None
    End If

    If cht.HasAxis(xlValue) Then
        If nextH Then
            ApplyAxisGridlines cht.Axes(xlValue)
        Else
            cht.Axes(xlValue).HasMajorGridlines = False
        End If
    End If

    If cht.HasAxis(xlCategory) Then
        If nextV Then
            ApplyAxisGridlines cht.Axes(xlCategory)
        Else
            cht.Axes(xlCategory).HasMajorGridlines = False
        End If
    End If
End Sub

Private Sub ApplyAxisGridlines(ax As Axis)
    If Not ax.HasMajorGridlines Then ax.HasMajorGridlines = True
    With ax.MajorGridlines.Format.Line
        .Visible = msoTrue
        .weight = gridlineWeight
        .DashStyle = msoLineSolid
        .ForeColor.rgb = colorNeutral2
    End With
End Sub


' ============================================================
'   REMOVE LEGEND AND RESIZE
' ============================================================
' Deletes the chart legend and resizes the plot area to standard web dimensions.
' Operates in-place — intended for iterative adjustment after chart creation.

Private Sub BuildRemoveLegendResize()
    If ActiveChart Is Nothing Then
        MsgNoActiveChart
        Exit Sub
    End If
    RemoveLegendAndResize ActiveChart
End Sub

Private Sub RemoveLegendAndResize(cht As Chart)
    If cht.HasLegend Then cht.Legend.Delete

    cht.PlotArea.Select
    Selection.Height = removeLegend_webHeight
    Selection.Top = removeLegend_webTop
    Selection.Width = removeLegend_webWidth
    Selection.Left = removeLegend_webLeft
End Sub

Sub RemoveLegendResizeButton()
    BuildRemoveLegendResize
End Sub


' ============================================================
'   RESET TO GREY
' ============================================================
' Grays out all series on a chart (line + fill).
' StartWithGray is the ribbon entry point: duplicates the source chart, then
' applies colorNeutral1. GrayOutChart is parameterised for potential pipeline use.

Public Sub GrayOutChart(Optional ByVal cht As Chart = Nothing, _
                        Optional ByVal duplicateChart As Boolean = True, _
                        Optional ByVal grayColor As Long = 0)
    On Error GoTo CleanFail

    Dim targetChart As Chart
    Dim i As Long, n As Long

    If cht Is Nothing Then
        If ActiveChart Is Nothing Then
            MsgNoActiveChart
            Exit Sub
        End If
        Set targetChart = ActiveChart
    Else
        Set targetChart = cht
    End If

    If grayColor = 0 Then grayColor = colorNeutral2

    If MsgGrayOutConfirm(duplicateChart) <> vbOK Then Exit Sub

    If duplicateChart Then
        targetChart.Parent.Duplicate.Select
        Set targetChart = ActiveChart
        If targetChart Is Nothing Then
            MsgCouldNotResolveDuplicate
            Exit Sub
        End If
    End If

    n = targetChart.SeriesCollection.Count
    For i = 1 To n
        With targetChart.SeriesCollection(i).Format
            With .Line
                .Visible = msoTrue
                .ForeColor.rgb = grayColor
            End With
            With .Fill
                .Visible = msoTrue
                .ForeColor.rgb = grayColor
                .Solid
            End With
        End With
    Next i

    Exit Sub

CleanFail:
    MsgError "GrayOutChart"
End Sub

Public Sub StartWithGray()
    GrayOutChart cht:=Nothing, duplicateChart:=True, grayColor:=colorNeutral1
End Sub
