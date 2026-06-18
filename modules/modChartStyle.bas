Attribute VB_Name = "modChartStyle"
'==== Module: modChartStyle ====
' Post-creation styling tools that re-apply the house style to an existing
' chart. Unlike the in-place toggles in modChartToggles, these (re)build chart
' chrome and may duplicate the source chart.
'
' Tools
' -----
'   LabelLastPointButton  - duplicates the chart and adds series-name labels to
'                           the final data point of each series (line charts get
'                           a narrowed plot area); preserves the original
'   ApplyChartStyle       - chart-type-agnostic styler: applies the house style
'                           in place, to the extent each chart type allows
'
' IsPieChartType is shared with modChartToggles, where it lives as a Public
' function; it is called here unqualified (VBA flat namespace).
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
'   PIE          - pie/donut: square-plot sizing + per-slice colours
'   AXIS_FILL    - bar/column/area: gridlines + axis labels + per-series FILL
'   LINE_SCATTER - line/scatter: axis labels + per-series LINE colour
'   OTHER        - radar/3D/surface/stock/treemap/combo/future: chrome only,
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
    ' source text boxes) only works reliably on the ACTIVE chart - otherwise
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

        Case Else   ' OTHER - safe chrome + canvas size, no plot geometry
            ApplyChartChrome cht
            InsertLogo cht, silent:=True
            InsertSource cht, silent:=True
            FormatTitle cht
            If cht.chartType = xlTreemap Or cht.chartType = xlSunburst Then
                ' Treemap/sunburst are a single series of points, so colour them
                ' per-point from the palette (matching the dedicated builder)
                ' rather than the per-series best-guess below. Other chartex types
                ' (waterfall, funnel, box & whisker, histogram) have real series and
                ' fall through to the per-series branch. Point-count read is guarded
                ' in case Excel rejects per-point colouring.
                Dim nTiles As Long
                On Error Resume Next
                nTiles = cht.SeriesCollection(1).Points.Count
                On Error GoTo CleanFail
                ' silent: this bucket degrades gracefully, so a chart that
                ' rejects per-point colouring must not surface a message.
                If nTiles > 0 Then ApplySliceColors cht, nTiles, silent:=True
            Else
                ' Best-guess colour; silent because some modern types (sunburst,
                ' funnel) don't accept standard per-series colouring.
                FormatSeriesColors cht, GetStyleColorMode(cht.chartType), silent:=True
            End If
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
' for axis-based charts - it would be wrong for pie/donut and other layouts.
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
' swallowed - if activation fails, the shape steps simply fall back to skipping.
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
