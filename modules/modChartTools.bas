Attribute VB_Name = "modChartTools"
'==== Module: modChartTools ====
' Post-creation chart utilities triggered from the Customisation and FillActions
' ribbon groups. All tools operate on an already-formatted chart.
'
' Tools
' -----
'   LabelLastPointButton  — adds series-name labels to the final data point of each
'                           series; duplicates the chart first
'   ToggleGridlines       — cycles major gridlines: None -> Horizontal -> Vertical -> Both
'   ToggleLegendButton       — toggles legend visibility and resizes the plot area;
'                           pie/donut use square plot area constants; operates in-place
'
' Duplication behaviour
' ---------------------
'   LabelLastPoint duplicates the source chart by default so the original is preserved.
'   ToggleGridlines and ToggleLegendButton operate in-place on the active chart —
'   they are intended for iterative adjustment, not one-shot creation.
Option Explicit


' ============================================================
'   LABEL LAST POINT
' ============================================================

Private Sub BuildLabelLastPoint()
    On Error GoTo CleanFail
    AppFast

    Dim ipts As Long
    Dim Npts As Long
    Dim bLabeled As Boolean
    Dim cht As Chart
    Dim srs As Series
    Dim shp As Shape
    Dim iColor As Long
    Dim lbl As DataLabel

    If ActiveChart Is Nothing Then
        MsgNoActiveChart
        GoTo CleanExit
    End If

    ' Duplicate the chart and capture new chart reference directly (no Select required)
    Dim dupShp As Shape
    Set dupShp = ActiveChart.Parent.Duplicate
    Set cht = dupShp.Chart

    ' Narrow plot area only on line charts to make room for end labels
    With cht.PlotArea
        If cht.chartType = xlLine Then
            .Width = chartWidth - labelLastPointPlotWidthInset
        End If
        .Left = 0
    End With

    ' Nudge Y-axis label box upward when legend is present
    If cht.hasLegend Then
        For Each shp In cht.Shapes
            If shp.name = "YAxisLabelBox" Then
                shp.IncrementTop labelLastPointTitleNudge
            End If
        Next shp
    End If

    ' Adjust plot area dimensions using direct object references (no Select required)
    Dim plHeight As Double
    Dim plWidth As Double
    With cht.PlotArea
        plHeight = .Height
        plWidth = .Width
        .Top = labelLastPointPlotTop
        .Width = plWidth * labelLastPointPlotWidthRatio
        .Height = plHeight
    End With

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
CleanExit:
    AppRestore
    Exit Sub

CleanFail:
    AppRestore
    MsgError "BuildLabelLastPoint"
End Sub

Sub LabelLastPointButton()
    BuildLabelLastPoint
End Sub


' ============================================================
'   TOGGLE GRIDLINES
' ============================================================
' Cycles major gridlines through four states in sequence:
'   None -> Horizontal only -> Vertical only -> Both -> None
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
        nextH = True:  nextV = False        ' None -> Horizontal only
    ElseIf hasH And Not hasV Then
        nextH = False: nextV = True         ' Horizontal -> Vertical only
    ElseIf Not hasH And hasV Then
        nextH = True:  nextV = True         ' Vertical -> Both
    Else
        nextH = False: nextV = False        ' Both -> None
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
    If cht.SeriesCollection.Count <= 1 Or cht.chartType = xlTreemap Then
        MsgLegendNotApplicable
        Exit Sub
    End If

    Dim addLegend As Boolean
    addLegend = Not cht.hasLegend

    If IsPieChartType(cht.chartType) Then
        ToggleLegendPie cht, addLegend
    Else
        ToggleLegendStandard cht, addLegend
    End If
End Sub

Private Function IsPieChartType(ByVal ct As Long) As Boolean
    IsPieChartType = (ct = xlPie Or ct = xlDoughnut Or _
                      ct = xlPieEx Or ct = xlDoughnutExploded)
End Function

' True for chart types that do not support xlLabelPositionOutsideEnd
' (stacked variants). Labels on these can only be centered, so the
' OUTSIDE state is skipped and contrast colouring is always applied.
Private Function IsCenterOnlyLabelType(ByVal ct As Long) As Boolean
    Select Case ct
        Case xlColumnStacked, xlColumnStacked100, _
             xlBarStacked, xlBarStacked100, _
             xlAreaStacked, xlAreaStacked100
            IsCenterOnlyLabelType = True
    End Select
End Function

Private Sub ToggleLegendStandard(cht As Chart, ByVal addLegend As Boolean)
    On Error GoTo CleanFail

    If addLegend Then
        cht.hasLegend = True
        With cht.Legend
            ' xlLegendPositionBottom lays entries out in a horizontal row and lets Excel
            ' auto-size the legend to its content width. Do NOT set .Width afterwards — that
            ' would override the content fit and (at full width) was forcing items to stack.
            ' Set .Top to place it vertically and .Left to anchor it to the left edge
            ' (Excel centers it horizontally by default); the content-fitted width is preserved.
            .Position = xlLegendPositionBottom
            .Font.Color = legendFontColor
            .Font.Size = axisFontSize
            .Top = legendTop
            .Left = legendLeftPad
        End With
        ' Shift the y-axis title box down to sit below the legend (with-legend layout)
        MoveYAxisLabelBox cht, yAxisLabelTop
        With cht.PlotArea
            .Height = plotAreaHeight
            .Top = plotAreaTop
            .Width = plotAreaWidth
            .Left = plotAreaLeft
        End With
    Else
        cht.Legend.Delete
        ' Restore the y-axis title box to its no-legend position
        MoveYAxisLabelBox cht, yAxisLabelTop_noLegend
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

' Moves the y-axis title box ("YAxisLabelBox") to an absolute top position.
' Setting an absolute Top (not IncrementTop) keeps this idempotent across repeated toggles.
Private Sub MoveYAxisLabelBox(cht As Chart, ByVal newTop As Single)
    Dim shp As Shape
    For Each shp In cht.Shapes
        If shp.name = "YAxisLabelBox" Then
            shp.Top = newTop
            Exit For
        End If
    Next shp
End Sub

Private Sub ToggleLegendPie(cht As Chart, ByVal addLegend As Boolean)
    On Error GoTo CleanFail

    Dim plotSize As Long
    Dim chtHeight As Double
    Dim chtWidth As Double

    If addLegend Then
        cht.hasLegend = True
        plotSize = pieplotAreaSize_legend
    Else
        cht.Legend.Delete
        plotSize = pieplotAreaSize_noLegend
    End If

    With cht.PlotArea
        .Width = plotSize
        .Height = plotSize
        .Left = pieplotAreaLeft
        .Top = pieplotAreaTop
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
'   None -> X only -> Y only -> Both -> None
' Uses TickLabelPosition to show/hide labels without removing the axis.
' Axes that do not exist on the chart are treated as "not visible" and skipped
' during assignment. Chart types with no axes (pie, donut) are a no-op.
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
        nextX = True:  nextY = False        ' None -> X only
    ElseIf hasX And Not hasY Then
        nextX = False: nextY = True         ' X only -> Y only
    ElseIf Not hasX And hasY Then
        nextX = True:  nextY = True         ' Y only -> Both
    Else
        nextX = False: nextY = False        ' Both -> None
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
'   APPLY CHART STYLE (chart-type agnostic)
' ============================================================
' Applies the house style to the SELECTED chart, in place (no retype, no
' duplication), to the extent possible for that chart type.
'
' Design principle: "skip what would be wrong", not "swallow exceptions".
' Every chart gets the universally-safe chrome (font, border, logo, source,
' title). Type-specific steps (plot geometry, axis formatting, series colour)
' are gated behind a positive classification so unsupported types degrade
' gracefully rather than producing wrong output (e.g. a rectangular pie or a
' uniformly-coloured pie series).
'
' Buckets (see ClassifyChart):
'   PIE          — pie/donut: square-plot sizing + per-slice colours
'   AXIS_FILL    — bar/column/area: gridlines + axis labels + per-series FILL
'   LINE_SCATTER — line/scatter: axis labels + per-series LINE colour
'   OTHER        — radar/3D/surface/stock/treemap/combo/future: chrome only,
'                  plus a best-guess per-series colour; no geometry/axis steps

Public Sub ApplyChartStyle()
    On Error GoTo CleanFail
    AppFast

    Dim cht As Chart
    Set cht = ResolveActiveChart()
    If cht Is Nothing Then
        MsgNoActiveChart
        GoTo CleanExit
    End If

    ' A ribbon click deselects the chart, but Chart.Shapes.Add* (logo, title and
    ' source text boxes) only works reliably on the ACTIVE chart — otherwise
    ' AddPicture raises 430 and AddTextbox silently fails. Activate it first.
    ActivateChart cht

    Dim bucket As String
    bucket = ClassifyChart(cht.chartType, cht)

    ' Geometry/layout must run BEFORE the logo/source/title boxes, because those
    ' are positioned at coordinates derived from chartWidth/chartHeight and the
    ' plot-area constants. The order within each bucket mirrors ApplyChartPipeline:
    ' size/layout -> logo -> source -> title -> axes -> colour.
    Select Case bucket
        Case "PIE"
            ' Pie path: square canvas + centred plot, per-slice colours.
            ' SetRoundChartSizeAndTitle also creates the title boxes and sizes the
            ' canvas, so no ApplyChartChrome/FormatTitle here.
            SetRoundChartSizeAndTitle cht, PieChartDefaults()
            InsertLogo cht, silent:=True
            InsertSource cht, silent:=True
            Dim nPts As Long
            On Error Resume Next
            nPts = cht.SeriesCollection(1).Points.Count
            On Error GoTo CleanFail
            If nPts > 0 Then ApplySliceColors cht, nPts

        Case "AXIS_FILL"
            ' OuterFormat sizes the canvas AND positions the plot area (incl.
            ' legend), so the title/source boxes sit clear of the plot.
            OuterFormat cht, DefaultChartDefaults()
            InsertLogo cht, silent:=True
            InsertSource cht, silent:=True
            FormatTitle cht
            FormatGridlines cht
            FormatXAxis cht
            FormatSeriesColors cht, "FILL", silent:=True

        Case "LINE_SCATTER"
            OuterFormat cht, DefaultChartDefaults()
            InsertLogo cht, silent:=True
            InsertSource cht, silent:=True
            FormatTitle cht
            FormatGridlines cht
            FormatXAxis cht
            FormatSeriesColors cht, "LINE", silent:=True

        Case Else   ' OTHER — safe chrome + canvas size, no plot geometry
            ApplyChartChrome cht
            InsertLogo cht, silent:=True
            InsertSource cht, silent:=True
            FormatTitle cht
            ' Best-guess colour; silent because some modern types (sunburst,
            ' treemap, funnel) don't accept standard per-series colouring.
            FormatSeriesColors cht, GetStyleColorMode(cht.chartType), silent:=True
    End Select

CleanExit:
    AppRestore
    Exit Sub
CleanFail:
    AppRestore
    MsgError "ApplyChartStyle"
End Sub

' Classifies a chart into the routing bucket used by ApplyChartStyle.
' Defaults to "OTHER" so any unrecognised type degrades gracefully.
Private Function ClassifyChart(ByVal ct As Long, cht As Chart) As String
    If IsPieChartType(ct) Then
        ClassifyChart = "PIE"
    ElseIf GetStyleColorMode(ct) = "LINE" Then
        ClassifyChart = "LINE_SCATTER"
    ElseIf cht.HasAxis(xlValue) Then
        ClassifyChart = "AXIS_FILL"
    Else
        ClassifyChart = "OTHER"
    End If
End Function

' The universally-safe + always-required subset of OuterFormat: font, border,
' and CANVAS SIZE. The title/subtitle/source/logo boxes are positioned at
' coordinates derived from chartWidth/chartHeight, so the canvas MUST be resized
' to those dimensions or the elements overlap the plot. Deliberately omits
' ApplyPlotAreaGeometry (the rectangular plot-area layout), which is applied only
' for axis-based charts — it would be wrong for pie/donut and other layouts.
Private Sub ApplyChartChrome(cht As Chart)
    On Error Resume Next
    cht.ChartArea.Font.name = fontPrimary
    cht.ChartArea.Border.LineStyle = xlNone
    ' Resize canvas to the configured dimensions (embedded charts only; a chart
    ' sheet has no settable size).
    If TypeName(cht.Parent) = "ChartObject" Then
        cht.Parent.Width = chartWidth
        cht.Parent.Height = chartHeight
    End If
    On Error GoTo 0
End Sub

' Activates a chart so Chart.Shapes.Add* operations work. Mirrors the pattern in
' modExport: for an embedded chart, activate the parent worksheet then the chart;
' for a chart sheet (Workbook parent), activate the chart directly. Errors are
' swallowed — if activation fails, the shape steps simply fall back to skipping.
Private Sub ActivateChart(cht As Chart)
    On Error Resume Next
    If TypeName(cht.Parent) = "ChartObject" Then
        cht.Parent.Parent.Activate   ' parent worksheet
        cht.Parent.Select            ' select the ChartObject
    End If
    cht.Activate
    On Error GoTo 0
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
'   TOGGLE DATA LABELS
' ============================================================
' Cycles data labels through three states: None -> Outside End -> Inside Centre -> None.
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

    ' Stacked types have no usable "outside end" position; skip the OUTSIDE state.
    Dim centerOnly As Boolean
    centerOnly = IsCenterOnlyLabelType(cht.chartType)

    ' Resolve target: single selected series, or Nothing for all series.
    ' Intentional: read user's current selection to determine single-series scope.
    Dim targetSrs As Series
    If TypeName(Selection) = "Series" Then
        Set targetSrs = Selection
    End If

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
    If centerOnly Then
        ' Two-state cycle: stacked types have no usable outside position.
        Select Case currentState
            Case "NONE": nextState = "INSIDE"
            Case Else:   nextState = "NONE"
        End Select
    Else
        Select Case currentState
            Case "NONE":    nextState = "OUTSIDE"
            Case "OUTSIDE": nextState = "INSIDE"
            Case Else:      nextState = "NONE"
        End Select
    End If

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
                        .Font.Size = axisFontSize
                        .Font.name = fontPrimary
                    End With
                    ' Pie/donut slices each have their own fill colour, so contrast
                    ' must be computed per point. All other types share one series
                    ' fill, so a single series-level contrast colour is correct.
                    If IsPieChartType(cht.chartType) Then
                        ApplyPerPointContrast srs
                    Else
                        srs.DataLabels.Font.Color = GetLabelContrastColor(srs)
                    End If
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

' Returns a label color (white or black) chosen for best contrast against the
' background the label sits on, based on WCAG relative luminance.
' For fill-based series (bars, columns, area) the background is the series fill.
' For line/scatter series the label floats against the plot background (white),
' so dark text is always used. Falls back to white on error.
Private Function GetLabelContrastColor(srs As Series) As Long
    On Error GoTo UseFallback

    ' Line/scatter: labels float on the plot background, not on a coloured fill.
    Dim ct As Long
    ct = srs.chartType
    Select Case ct
        Case xlLine, xlLineMarkers, xlLineStacked, xlLineMarkersStacked, _
             xlLineStacked100, xlLineMarkersStacked100, _
             xlXYScatter, xlXYScatterLines, xlXYScatterLinesNoMarkers, _
             xlXYScatterSmooth, xlXYScatterSmoothNoMarkers
            GetLabelContrastColor = colorBrand3
            Exit Function
    End Select

    GetLabelContrastColor = ContrastColorForFill(srs.Format.Fill.ForeColor.RGB)
    Exit Function

UseFallback:
    GetLabelContrastColor = colorWhite
End Function

' Colours each point's data label against that point's own fill. Used for
' pie/donut, where every slice has a distinct colour, so a single series-level
' contrast colour (as GetLabelContrastColor returns) would be wrong for all but
' one slice. Per-point errors are skipped so one bad slice can't abort the rest.
Private Sub ApplyPerPointContrast(srs As Series)
    Dim p As Long
    Dim nPts As Long

    On Error Resume Next
    nPts = srs.Points.Count
    On Error GoTo 0
    If nPts = 0 Then Exit Sub

    For p = 1 To nPts
        On Error Resume Next
        srs.Points(p).DataLabel.Font.Color = _
            ContrastColorForFill(srs.Points(p).Format.Fill.ForeColor.RGB)
        On Error GoTo 0
    Next p
End Sub

' Returns white or dark brand text for best contrast against the given fill
' colour, using WCAG relative luminance: white text on dark fills, dark text
' on light fills.
Private Function ContrastColorForFill(ByVal fillRGB As Long) As Long
    If RelativeLuminance(fillRGB) < wcagLuminanceThreshold Then
        ContrastColorForFill = colorWhite
    Else
        ContrastColorForFill = colorBrand3
    End If
End Function

' Calculates WCAG relative luminance of an RGB color.
' Input: clr as Long (Excel RGB format: R + G*256 + B*65536)
' Returns: Double in range [0, 1], where 0 is black and 1 is white
Private Function RelativeLuminance(ByVal clr As Long) As Double
    Dim R As Double, G As Double, B As Double
    Dim Rs As Double, Gs As Double, Bs As Double

    ' Extract 8-bit RGB components from Long
    R = clr Mod 256
    G = (clr \ 256) Mod 256
    B = (clr \ 65536) Mod 256

    ' Normalize to 0–1
    Rs = R / 255#
    Gs = G / 255#
    Bs = B / 255#

    ' Convert sRGB to linear RGB
    If Rs <= 0.03928 Then
        Rs = Rs / 12.92
    Else
        Rs = ((Rs + 0.055) / 1.055) ^ 2.4
    End If

    If Gs <= 0.03928 Then
        Gs = Gs / 12.92
    Else
        Gs = ((Gs + 0.055) / 1.055) ^ 2.4
    End If

    If Bs <= 0.03928 Then
        Bs = Bs / 12.92
    Else
        Bs = ((Bs + 0.055) / 1.055) ^ 2.4
    End If

    ' WCAG luminance formula
    RelativeLuminance = 0.2126 * Rs + 0.7152 * Gs + 0.0722 * Bs
End Function
