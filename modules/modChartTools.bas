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
'   ToggleLegendButton       — toggles legend visibility and resizes the plot area;
'                           pie/donut use square plot area constants; operates in-place
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
    If ActiveChart.chartType = xlLine Then
        Selection.Width = chartWidth - labelLastPointPlotWidthInset
    End If
    Selection.Left = 0

    Set cht = ActiveChart

    ' Nudge Y-axis label box upward when legend is present
    If cht.hasLegend Then
        For Each shp In cht.Shapes
            If shp.name = "YAxisLabelBox" Then
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
    Selection.Width = plWidth * labelLastPointPlotWidthRatio
    Selection.Height = plHeight

    ' Remove legend (labels replace it)
    If cht.hasLegend Then
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

                        Select Case srs.chartType
                            Case xlLine, xlLineStacked, xlLineStacked100, xlLineMarkers, xlLineMarkersStacked, xlLineMarkersStacked100
                                lbl.Position = xlLabelPositionRight
                                iColor = .Format.Line.ForeColor.RGB
                            Case xlXYScatter, xlXYScatterLines, xlXYScatterLinesNoMarkers, xlXYScatterSmooth, xlXYScatterSmoothNoMarkers
                                lbl.Position = xlLabelPositionRight
                                iColor = .MarkerBackgroundColor
                            Case xlColumnClustered, xlBarClustered
                                lbl.Position = xlLabelPositionOutsideEnd
                                iColor = .Format.Fill.ForeColor.RGB
                            Case xlColumnStacked, xlColumnStacked100, xlBarStacked, xlBarStacked100, xlArea, xlAreaStacked, xlAreaStacked100
                                lbl.Position = xlLabelPositionCenter
                                iColor = .Format.Fill.ForeColor.RGB
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

    ' Read gridline state via helpers that work even when the axis has been removed.
    hasH = GetGridlineState(cht, xlValue)
    hasV = GetGridlineState(cht, xlCategory)

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

    ' Apply gridlines only to currently visible axes; remove from any axis (even hidden ones).
    If nextH And cht.HasAxis(xlValue) Then
        ApplyAxisGridlines cht.Axes(xlValue)
    End If
    If Not nextH Then ClearGridlinesSafe cht, xlValue

    If nextV And cht.HasAxis(xlCategory) Then
        ApplyAxisGridlines cht.Axes(xlCategory)
    End If
    If Not nextV Then ClearGridlinesSafe cht, xlCategory
End Sub

' Returns True if the axis has major gridlines, temporarily re-enabling the axis
' if it has been removed so the property can be read reliably.
Private Function GetGridlineState(cht As Chart, ByVal axisType As Long) As Boolean
    Dim wasVisible As Boolean
    wasVisible = cht.HasAxis(axisType, xlPrimary)

    If Not wasVisible Then
        On Error Resume Next
        cht.HasAxis(axisType, xlPrimary) = True
        If Err.Number <> 0 Then Err.Clear: Exit Function   ' axis not available
        On Error GoTo 0
    End If

    On Error Resume Next
    If cht.HasAxis(axisType) Then GetGridlineState = cht.Axes(axisType).HasMajorGridlines
    On Error GoTo 0

    If Not wasVisible Then
        On Error Resume Next
        cht.HasAxis(axisType, xlPrimary) = False
        On Error GoTo 0
    End If
End Function

' Removes major gridlines from the axis, temporarily re-enabling it if it has been removed.
Private Sub ClearGridlinesSafe(cht As Chart, ByVal axisType As Long)
    Dim wasVisible As Boolean
    wasVisible = cht.HasAxis(axisType, xlPrimary)

    If Not wasVisible Then
        On Error Resume Next
        cht.HasAxis(axisType, xlPrimary) = True
        If Err.Number <> 0 Then Err.Clear: Exit Sub        ' axis not available
        On Error GoTo 0
    End If

    On Error Resume Next
    If cht.HasAxis(axisType) Then cht.Axes(axisType).HasMajorGridlines = False
    On Error GoTo 0

    If Not wasVisible Then
        On Error Resume Next
        cht.HasAxis(axisType, xlPrimary) = False
        On Error GoTo 0
    End If
End Sub

Private Sub ApplyAxisGridlines(ax As Axis)
    If Not ax.HasMajorGridlines Then ax.HasMajorGridlines = True
    With ax.MajorGridlines.Format.Line
        .Visible = msoTrue
        .Weight = gridlineWeight
        .DashStyle = msoLineSolid
        .ForeColor.RGB = colorNeutral2
    End With
End Sub


' ============================================================
'   TOGGLE AxES
' ============================================================
' Cycles axis visibility through four states in sequence:
'   None → Y axis only → X axis only → Both → None
' Operates in-place on the active chart (no duplication).

Public Sub ToggleAxes()
    If ActiveChart Is Nothing Then
        MsgNoActiveChart
        Exit Sub
    End If

    Dim cht As Chart
    Set cht = ActiveChart

    Dim hasY As Boolean     ' value axis (Y)
    Dim hasX As Boolean     ' category axis (X)

    hasY = cht.HasAxis(xlValue)
    hasX = cht.HasAxis(xlCategory)

    Dim nextY As Boolean
    Dim nextX As Boolean

    If Not hasY And Not hasX Then
        nextY = True:  nextX = False        ' None → Y only
    ElseIf hasY And Not hasX Then
        nextY = False: nextX = True         ' Y only → X only
    ElseIf Not hasY And hasX Then
        nextY = True:  nextX = True         ' X only → Both
    Else
        nextY = False: nextX = False        ' Both → None
    End If

    cht.HasAxis(xlValue, xlPrimary) = nextY
    cht.HasAxis(xlCategory, xlPrimary) = nextX

    If nextY Then ApplyValueAxisStyle cht
    If nextX Then ApplyCategoryAxisStyle cht
End Sub

Private Sub ApplyValueAxisStyle(cht As Chart)
    If Not cht.HasAxis(xlValue) Then Exit Sub
    With cht.Axes(xlValue)
        .Format.Line.Visible = msoFalse
        .TickLabels.Font.Size = axisFontSize
        .TickLabels.Font.Color = colorBrand3
    End With
End Sub

Private Sub ApplyCategoryAxisStyle(cht As Chart)
    If Not cht.HasAxis(xlCategory) Then Exit Sub
    cht.Axes(xlCategory).TickLabels.Font.Size = axisFontSize
    cht.Axes(xlCategory, xlPrimary).TickLabels.Font.Color = colorBrand3
    cht.Axes(xlCategory).Select
    With Selection.Format.Line
        .Visible = msoTrue
        .ForeColor.RGB = colorBrand3
        .ForeColor.TintAndShade = 0
        .ForeColor.Brightness = 0
        .Weight = axisLineWeight
    End With
End Sub


' ============================================================
'   TOGGLE LEGEND
' ============================================================
' Toggles legend visibility and resizes the plot area to match.
' Pie/donut:       uses square plot area constants from modConfig.
' Standard charts: uses remove-legend / with-legend constants from modConfig.
' Single-series or treemap charts: informational message, no change.
' Operates in-place on the active chart (no duplication).

Public Sub ToggleLegend()
    If ActiveChart Is Nothing Then
        MsgNoActiveChart
        Exit Sub
    End If

    Dim cht As Chart
    Set cht = ActiveChart

    ' Single-series: legend is redundant. Treemap uses tile labels instead.
    If cht.SeriesCollection.Count <= 1 Or cht.ChartType = xlTreemap Then
        MsgLegendNotApplicable
        Exit Sub
    End If

    Dim addLegend As Boolean
    addLegend = Not cht.HasLegend

    If IsPieChartType(cht.ChartType) Then
        ToggleLegendPie cht, addLegend
    Else
        ToggleLegendStandard cht, addLegend
    End If
End Sub

Private Function IsPieChartType(ByVal ct As Long) As Boolean
    IsPieChartType = (ct = xlPie Or ct = xlDoughnut Or _
                      ct = xlPie3D Or ct = xlDoughnutExploded)
End Function

Private Sub ToggleLegendStandard(cht As Chart, ByVal addLegend As Boolean)
    On Error GoTo CleanFail

    If addLegend Then
        cht.HasLegend = True
        With cht.Legend
            .Position = xlLegendPositionTop
            .Left = legendLeftPad
            .Font.Color = legendFontColor
            .Font.Size = axisFontSize
        End With
        With cht.PlotArea
            .Height = PlotAreaHeight
            .Top = PlotAreaTop
            .Width = plotAreaWidth
            .Left = plotAreaLeft
        End With
    Else
        cht.Legend.Delete
        With cht.PlotArea
            .Height = removelegendHeight
            .Top = removelegendTop
            .Width = removeLegend_Width
            .Left = removeLegend_Left
        End With
    End If

    Exit Sub
CleanFail:
    MsgError "ToggleLegendStandard"
End Sub

Private Sub ToggleLegendPie(cht As Chart, ByVal addLegend As Boolean)
    On Error GoTo CleanFail

    Dim plotSize As Long
    Dim chtHeight As Double
    Dim chtWidth As Double

    If addLegend Then
        cht.HasLegend = True
        plotSize = piePlotAreaSize_legend
    Else
        cht.Legend.Delete
        plotSize = piePlotAreaSize_noLegend
    End If

    With cht.PlotArea
        .Width = plotSize
        .Height = plotSize
        .Left = piePlotAreaLeft
        .Top = piePlotAreaTop
    End With

    chtHeight = cht.ChartArea.Height
    chtWidth = cht.ChartArea.Width
    cht.PlotArea.Top = (chtHeight - plotSize) * piePlotTopRatio
    cht.PlotArea.Left = (chtWidth - plotSize) / 2

    If addLegend Then
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
    MsgError "ToggleLegendPie"
End Sub

Sub ToggleLegendButton()
    ToggleLegend
End Sub


' ============================================================
'   TOGGLE AXIS LABELS
' ============================================================
' Cycles axis tick-label visibility through four states in sequence:
'   None → X only → Y only → Both → None
' Uses TickLabelPosition to show/hide labels without removing the axis.
' Axes that do not exist (removed via ToggleAxes) are treated as "not visible"
' and skipped during assignment. Chart types with no axes (pie, donut) are a no-op.
' Operates in-place on the active chart (no duplication).

Public Sub ToggleAxisLabels()
    If ActiveChart Is Nothing Then
        MsgNoActiveChart
        Exit Sub
    End If

    Dim cht As Chart
    Set cht = ActiveChart

    Dim hasX As Boolean   ' category axis labels visible
    Dim hasY As Boolean   ' value axis labels visible

    hasX = GetAxisLabelState(cht, xlCategory)
    hasY = GetAxisLabelState(cht, xlValue)

    Dim nextX As Boolean
    Dim nextY As Boolean

    If Not hasX And Not hasY Then
        nextX = True:  nextY = False        ' None → X only
    ElseIf hasX And Not hasY Then
        nextX = False: nextY = True         ' X only → Y only
    ElseIf Not hasX And hasY Then
        nextX = True:  nextY = True         ' Y only → Both
    Else
        nextX = False: nextY = False        ' Both → None
    End If

    If cht.HasAxis(xlCategory) Then SetAxisLabelState cht.Axes(xlCategory), nextX
    If cht.HasAxis(xlValue) Then SetAxisLabelState cht.Axes(xlValue), nextY
End Sub

' Returns True if the axis exists and has visible tick labels.
Private Function GetAxisLabelState(cht As Chart, ByVal axisType As Long) As Boolean
    If Not cht.HasAxis(axisType) Then Exit Function
    On Error Resume Next
    GetAxisLabelState = (cht.Axes(axisType).TickLabelPosition <> xlTickLabelPositionNone)
    On Error GoTo 0
End Function

' Shows or hides tick labels on an axis; applies brand styling when showing.
Private Sub SetAxisLabelState(ax As Axis, ByVal show As Boolean)
    On Error Resume Next
    If show Then
        ax.TickLabelPosition = xlTickLabelPositionNextToAxis
        ax.TickLabels.Font.Size = axisFontSize
        ax.TickLabels.Font.Color = axisFontColor
    Else
        ax.TickLabelPosition = xlTickLabelPositionNone
    End If
    On Error GoTo 0
End Sub

Sub ToggleAxisLabelsButton()
    ToggleAxisLabels
End Sub


' ============================================================
'   APPLY CHART STYLE (generic)
' ============================================================
' Applies brand formatting to the active chart regardless of type.
' Operates in-place — no duplication. Steps that require axes are
' skipped when the chart type does not have them (e.g. pie/donut).

Public Sub ApplyChartStyle()
    If ActiveChart Is Nothing Then
        MsgNoActiveChart
        Exit Sub
    End If

    Dim cht As Chart
    Set cht = ActiveChart

    OuterFormat cht, DefaultChartDefaults()
    FormatXAxisTitle cht
    InsertLogo cht
    InsertSource cht
    FormatTitle cht
    If cht.HasAxis(xlValue) Then FormatGridlines cht
    FormatXAxis cht
    FormatSeriesColors cht, GetStyleColorMode(cht.chartType)
End Sub

Private Function GetStyleColorMode(ByVal ct As Long) As String
    Select Case ct
        Case xlLine, xlLineMarkers, xlLineStacked, xlLineMarkersStacked, _
             xlLineStacked100, xlLineMarkersStacked100, _
             xlXYScatter, xlXYScatterLines, xlXYScatterLinesNoMarkers, _
             xlXYScatterSmooth, xlXYScatterSmoothNoMarkers
            GetStyleColorMode = "LINE"
        Case Else
            GetStyleColorMode = "FILL"
    End Select
End Function


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
                .ForeColor.RGB = grayColor
            End With
            With .Fill
                .Visible = msoTrue
                .ForeColor.RGB = grayColor
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


' ============================================================
'   TOGGLE DATA LABELS
' ============================================================
' Cycles data labels through three states: None → Outside End → Inside Centre → None.
' Operates in-place on the active chart.
' Scope: if a specific series is selected, only that series is affected;
'        otherwise all series in the chart are affected.

Public Sub ToggleDataLabels()
    If ActiveChart Is Nothing Then
        MsgNoActiveChart
        Exit Sub
    End If

    Dim cht As Chart
    Set cht = ActiveChart

    ' Resolve target: single selected series, or Nothing for all series.
    Dim targetSrs As Series
    On Error Resume Next
    Set targetSrs = Selection
    On Error GoTo 0

    ' Detect current state from the first (or only) target series.
    Dim firstSrs As Series
    If Not targetSrs Is Nothing Then
        Set firstSrs = targetSrs
    ElseIf cht.SeriesCollection.Count > 0 Then
        Set firstSrs = cht.SeriesCollection(1)
    End If

    Dim currentState As String
    currentState = "NONE"
    If Not firstSrs Is Nothing Then
        If firstSrs.HasDataLabels Then
            Dim pos As Long
            On Error Resume Next
            pos = firstSrs.DataLabels.Position
            If Err.Number <> 0 Then
                currentState = "OTHER"   ' can't read position — treat as non-standard
            Else
                On Error GoTo 0
                Select Case pos
                    Case xlLabelPositionOutsideEnd:  currentState = "OUTSIDE"
                    Case xlLabelPositionCenter:      currentState = "INSIDE"
                    Case Else:                       currentState = "OTHER"
                End Select
            End If
            On Error GoTo 0
        End If
    End If

    ' Advance to next state.
    Dim nextState As String
    Select Case currentState
        Case "NONE":    nextState = "OUTSIDE"
        Case "OUTSIDE": nextState = "INSIDE"
        Case Else:      nextState = "NONE"
    End Select

    ' Apply to target series or all series.
    Dim i As Long
    Dim n As Long
    If Not targetSrs Is Nothing Then
        n = 1
    Else
        n = cht.SeriesCollection.Count
    End If

    For i = 1 To n
        Dim srs As Series
        If Not targetSrs Is Nothing Then
            Set srs = targetSrs
        Else
            Set srs = cht.SeriesCollection(i)
        End If

        Select Case nextState
            Case "NONE"
                srs.HasDataLabels = False

            Case "OUTSIDE"
                ' Try preferred position; fall back to center if unsupported.
                If Not TrySetLabelPosition(srs, xlLabelPositionOutsideEnd) Then
                    TrySetLabelPosition srs, xlLabelPositionCenter
                End If
                If srs.HasDataLabels Then
                    With srs.DataLabels
                        .Font.Color = colorBrand3
                        .Font.Size = axisFontSize
                        .Font.name = fontPrimary
                    End With
                End If

            Case "INSIDE"
                ' Try preferred position; fall back to outside end if unsupported.
                If Not TrySetLabelPosition(srs, xlLabelPositionCenter) Then
                    TrySetLabelPosition srs, xlLabelPositionOutsideEnd
                End If
                If srs.HasDataLabels Then
                    With srs.DataLabels
                        .Font.Color = GetLabelContrastColor(srs)
                        .Font.Size = axisFontSize
                        .Font.name = fontPrimary
                    End With
                End If
        End Select
    Next i
End Sub

' Attempts to apply data labels to a series with the given position.
' Returns True on success, False if the chart type does not support the position.
Private Function TrySetLabelPosition(srs As Series, ByVal pos As Long) As Boolean
    On Error GoTo Fail
    srs.ApplyDataLabels
    srs.DataLabels.Position = pos
    TrySetLabelPosition = True
    Exit Function
Fail:
End Function

' ============================================================
'   TOGGLE CHART VARIANT
' ============================================================
' Switches the active chart between its default and alternative type:
'   Stacked Bar    ↔  100% Stacked Bar
'   Stacked Column ↔  100% Stacked Column
'   Pie            ↔  Donut
'   Line           ↔  Line with Markers
'   Stacked Area   ↔  100% Stacked Area
' Operates in-place. Does nothing for unsupported chart types.

Public Sub ToggleChartVariant()
    If ActiveChart Is Nothing Then
        MsgNoActiveChart
        Exit Sub
    End If

    Dim altType As Long
    altType = GetAlternativeChartType(ActiveChart.chartType)

    If altType = -1 Then Exit Sub   ' unsupported type — do nothing

    ActiveChart.chartType = altType
End Sub

Private Function GetAlternativeChartType(ByVal ct As Long) As Long
    Select Case ct
        Case xlBarStacked:       GetAlternativeChartType = xlBarStacked100
        Case xlBarStacked100:    GetAlternativeChartType = xlBarStacked
        Case xlColumnStacked:    GetAlternativeChartType = xlColumnStacked100
        Case xlColumnStacked100: GetAlternativeChartType = xlColumnStacked
        Case xlPie:              GetAlternativeChartType = xlDoughnut
        Case xlDoughnut:         GetAlternativeChartType = xlPie
        Case xlLine:             GetAlternativeChartType = xlLineMarkers
        Case xlLineMarkers:      GetAlternativeChartType = xlLine
        Case xlAreaStacked:      GetAlternativeChartType = xlAreaStacked100
        Case xlAreaStacked100:   GetAlternativeChartType = xlAreaStacked
        Case Else:               GetAlternativeChartType = -1
    End Select
End Function


' Returns a label color (colorBrand3 dark or colorBrand4 light) chosen for best
' contrast against the series fill color. Falls back to colorBrand4 on any error.
Private Function GetLabelContrastColor(srs As Series) As Long
    On Error GoTo UseFallback

    Dim fillRGB As Long
    fillRGB = srs.Format.Fill.ForeColor.RGB

    ' Excel stores RGB as R + G*256 + B*65536
    Dim r As Long, g As Long, b As Long
    r = fillRGB And &HFF
    g = (fillRGB \ 256) And &HFF
    b = (fillRGB \ 65536) And &HFF

    ' Perceived luminance (ITU-R BT.601)
    Dim lum As Double
    lum = 0.299 * r + 0.587 * g + 0.114 * b

    ' Dark fill → use light label (colorBrand4); light fill → use dark label (colorBrand3)
    If lum < 128 Then
        GetLabelContrastColor = colorBrand4
    Else
        GetLabelContrastColor = colorBrand3
    End If
    Exit Function

UseFallback:
    GetLabelContrastColor = colorBrand4
End Function
