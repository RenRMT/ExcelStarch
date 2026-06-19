Attribute VB_Name = "modEngineStyle"
'==== Module: modEngineStyle ====
' Post-creation styling tools that re-apply the house style to an existing
' chart. Unlike the in-place toggles in modEngineToggles, these (re)build chart
' chrome and may duplicate the source chart.
'
' Tools
' -----
'   AnnotateButton        - adds an editable annotation text box to the active
'                           chart. With nothing (or a background element)
'                           selected the box lands at the plot-area centre; with
'                           a data point selected it sits near that point; with a
'                           whole series selected it sits near the series' last
'                           point. Chart-type agnostic - no per-type layout.
'   ApplyChartStyle       - chart-type-agnostic styler: applies the house style
'                           in place, to the extent each chart type allows
'
' IsPieChartType is shared with modEngineToggles, where it lives as a Public
' function; it is called here unqualified (VBA flat namespace).
Option Explicit


' ============================================================
'   ANNOTATION
' ============================================================
' Adds an editable annotation text box to the active chart. Placement depends
' on the current selection:
'   - a single data Point        -> near that point
'   - a whole Series             -> near the series' last point
'   - nothing / background        -> centre of the plot area
' The box is additive (named "AnnotationBox<n>"), so repeated runs stack rather
' than overwrite, and the user can freely retype the default text.

Public Sub AnnotateButton()
    BuildAnnotation
End Sub

Private Sub BuildAnnotation()
    On Error GoTo CleanFail
    AppFast

    Dim cht As Chart
    Set cht = ResolveActiveChart()
    If cht Is Nothing Then
        MsgNoActiveChart
        GoTo CleanExit
    End If

    ' Chart.Shapes.AddTextbox only works reliably on the ACTIVE chart - a ribbon
    ' click deselects it, so re-activate first (mirrors ApplyChartStyle).
    ActivateChart cht

    Dim pt As Point
    Set pt = ResolveSelectedPoint()      ' Nothing => centre case

    Dim leftPos As Double, topPos As Double
    If pt Is Nothing Then
        GetPlotCenter cht, leftPos, topPos
    ElseIf Not TryGetPointAnchor(pt, leftPos, topPos) Then
        ' Unplotted point has no pixel anchor - fall back to centre.
        GetPlotCenter cht, leftPos, topPos
    End If

    AddAnnotationBox cht, leftPos, topPos

CleanExit:
    AppRestore
    Exit Sub

CleanFail:
    AppRestore
    MsgError "BuildAnnotation"
End Sub

' Returns the selected Point to anchor to, or Nothing for the plot-centre case.
' A selected Series resolves to its last point. Mirrors the Series/Point
' detection in modColorFill (GetFillTarget / IsSeriesOrPoint).
Private Function ResolveSelectedPoint() As Point
    If Selection Is Nothing Then Exit Function

    Dim pt As Point
    Dim srs As Series

    On Error Resume Next
    Set pt = Selection
    If Err.Number = 0 And Not pt Is Nothing Then
        Set ResolveSelectedPoint = pt
        Err.Clear
        On Error GoTo 0
        Exit Function
    End If
    Err.Clear

    Set srs = Selection
    If Err.Number = 0 And Not srs Is Nothing Then
        ' Series selected -> anchor on its last point.
        If srs.Points.Count > 0 Then Set ResolveSelectedPoint = srs.Points(srs.Points.Count)
    End If
    Err.Clear
    On Error GoTo 0
End Function

' Reads a data point's chart-relative pixel position by briefly borrowing its
' DataLabel (a Point exposes no .Left/.Top). Any label we create is removed
' again so the chart is left untouched. Returns False if the point is unplotted
' or otherwise has no readable position.
Private Function TryGetPointAnchor(pt As Point, ByRef leftPos As Double, ByRef topPos As Double) As Boolean
    Dim hadLabel As Boolean
    Dim lbl As DataLabel

    On Error GoTo Fail
    hadLabel = pt.HasDataLabel
    If Not hadLabel Then pt.HasDataLabel = True

    Set lbl = pt.DataLabel
    leftPos = lbl.Left + annotationOffsetX
    topPos = lbl.Top + annotationOffsetY

    If Not hadLabel Then pt.HasDataLabel = False   ' restore: we added it
    On Error GoTo 0
    TryGetPointAnchor = True
    Exit Function

Fail:
    ' Best-effort cleanup of a label we may have added before failing.
    On Error Resume Next
    If Not hadLabel Then pt.HasDataLabel = False
    On Error GoTo 0
    TryGetPointAnchor = False
End Function

' Top-left of an annotation box centred on the plotting region. Uses Inside*
' so the box centres on the actual plot, not the plot-area frame.
Private Sub GetPlotCenter(cht As Chart, ByRef leftPos As Double, ByRef topPos As Double)
    With cht.PlotArea
        leftPos = .InsideLeft + .InsideWidth / 2 - annotationBoxWidth / 2
        topPos = .InsideTop + .InsideHeight / 2 - annotationBoxHeight / 2
    End With
End Sub

' Creates the annotation text box at the given chart-relative position. The name
' is suffixed with the next free index so repeated runs stack instead of
' overwriting (so we deliberately do NOT SafeDeleteShape first).
Private Sub AddAnnotationBox(cht As Chart, ByVal leftPos As Double, ByVal topPos As Double)
    Dim shp As Shape
    Set shp = cht.Shapes.AddTextbox( _
                    Orientation:=msoTextOrientationHorizontal, _
                    Left:=leftPos, Top:=topPos, _
                    Width:=annotationBoxWidth, Height:=annotationBoxHeight)

    With shp
        .name = NextAnnotationName(cht)
        .TextFrame2.TextRange.Text = annotationDefaultText
        With .TextFrame2.TextRange.Font
            .Size = annotationFontSize
            .name = fontPrimary
            .Fill.ForeColor.RGB = annotationFontColor
            .Bold = msoFalse
        End With
    End With
End Sub

' Returns the next free "AnnotationBox<n>" name by scanning existing shapes.
Private Function NextAnnotationName(cht As Chart) As String
    Dim n As Long
    n = 0
    Dim shp As Shape
    For Each shp In cht.Shapes
        If Left$(shp.name, 13) = "AnnotationBox" Then n = n + 1
    Next shp
    NextAnnotationName = "AnnotationBox" & (n + 1)
End Function


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
