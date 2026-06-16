Attribute VB_Name = "modTreemapChrome"
'==== Module: modTreemapChrome ====
' Worksheet-shape chrome for treemap charts.
'
' Why this exists
' ---------------
' xlTreemap (a "chartex" type, like sunburst/waterfall/funnel/box-whisker) rejects
' cht.Shapes.AddTextbox / AddPicture with error 1004 - the chartex file schema has no
' userShapes slot, so a treemap cannot own the title/subtitle/logo/source overlay
' boxes that classic charts carry inside cht.Shapes. The chrome is therefore created
' on the HOST WORKSHEET (cht.Parent.Parent.Shapes), positioned over the chart, and
' grouped with the ChartObject so the group exports as one image.
'
' This is deliberately a separate pipeline from the classic in-chart chrome in
' modChartBuilder (which is left untouched). Treemaps have no value axis and no
' legend, so the chrome is six static shapes: a white 600x600 Canvas behind
' everything, then FigureBox, TitleBox, SubTitleBox, SourceBox and LogoImage on top
' of it - there is no YAxisLabelBox and no legend-toggle repositioning. The text
' boxes are transparent with no border; the Canvas supplies the white background.
'
' Coordinate model
' ----------------
' Classic chrome uses chart-relative coordinates on the 600x600 canvas. Worksheet
' shapes use absolute sheet coordinates, so every position is offset by the chart's
' position on the sheet (cht.Parent.Left / cht.Parent.Top). The chart is forced to
' chartWidth x chartHeight before chrome is built, so the same modConfig geometry
' constants apply.
'
' Naming
' ------
' Shape members are named "<ChartObjectName><suffix>" (e.g. "Chart 1_TitleBox") so
' that several treemaps on one sheet do not collide. The group is named
' "ESTreemapGroup_<ChartObjectName>" so it can be found again on re-run/export.
'
' Known limitations (prototype)
' -----------------------------
'   - Moving the chart after creation leaves the chrome behind once the group is
'     ungrouped (re-styling ungroups). Re-run the builder to realign.
'   - Re-running with the GROUP selected (rather than the chart) cannot retype the
'     chart; ActiveChart is Nothing and the caller shows MsgSelectRangeOrChart.
'     Click the chart itself, not the group.
'   - Export rasterises the group at ~screen resolution (see modExport), softer than
'     the classic Chart.Export path.
'
' Z-order note / future refinement
' --------------------------------
' The ChartObject is created by the shared GetTargetChart BEFORE this module runs, so
' the white Canvas (added here) lands in front of it and is pushed behind with
' ZOrder msoSendToBack. A cleaner design would create the Canvas FIRST and add the
' chart + chrome on top, making the z-order correct by construction - but that needs a
' treemap-specific creation flow rather than the shared GetTargetChart. Deferred to
' the eventual chartex pipeline generalisation.
Option Explicit

' Public so modExport can recognise treemap groups by name (single source of truth).
Public Const treemapGroupPrefix As String = "ESTreemapGroup_"


' ============================================================
'   PUBLIC ENTRY
' ============================================================

' Builds the chrome shapes on the host worksheet at the given canvas origin, and
' groups them with the ChartObject. Removes any prior chrome/group first so re-runs
' don't duplicate. baseLeft/baseTop are the canvas top-left (resolved by
' TreemapCanvasOrigin before the chart was repositioned). Caller has already shrunk
' the chart into the plot band and coloured the tiles.
Public Sub BuildTreemapChrome(cht As Chart, ByVal baseLeft As Double, ByVal baseTop As Double)
    On Error GoTo CleanFail

    Dim ws As Worksheet
    Set ws = HostSheet(cht)
    If ws Is Nothing Then
        MsgTreemapNeedsEmbedded
        Exit Sub
    End If

    ' Clear any chrome from a previous run before rebuilding.
    RemoveExistingTreemapChrome cht

    Dim baseName As String
    baseName = cht.Parent.name

    Dim chromeNames(1 To 6) As String
    Dim shp As Shape

    ' White canvas first so it sits at the back of the z-order; everything else
    ' (chart tiles + chrome) renders on top of it.
    Set shp = AddTreemapCanvas(ws, baseName, baseLeft, baseTop): chromeNames(1) = shp.name

    Set shp = AddTreemapFigureBox(ws, baseName, baseLeft, baseTop): chromeNames(2) = shp.name
    Set shp = AddTreemapTitleBox(ws, baseName, baseLeft, baseTop): chromeNames(3) = shp.name
    Set shp = AddTreemapSubtitleBox(ws, baseName, baseLeft, baseTop): chromeNames(4) = shp.name
    Set shp = AddTreemapSourceBox(ws, baseName, baseLeft, baseTop): chromeNames(5) = shp.name

    ' Logo may fail to decode; build the other shapes regardless.
    Set shp = AddTreemapLogo(ws, baseName, baseLeft, baseTop)
    If shp Is Nothing Then
        chromeNames(6) = vbNullString
    Else
        chromeNames(6) = shp.name
    End If

    GroupTreemapChrome ws, cht, chromeNames

    Exit Sub
CleanFail:
    MsgError "BuildTreemapChrome"
End Sub


' ============================================================
'   SHAPE BUILDERS (worksheet-targeted, offset by chart position)
' ============================================================
' Bodies mirror the classic Create*Box / InsertSource / InsertLogo helpers in
' modChartBuilder, but target ws.Shapes and add (baseLeft, baseTop) to every
' position. baseName is the ChartObject name, used to make member names unique.

' A borderless white 600x600 rectangle behind the chart and chrome, so the whole
' group exports on a solid white canvas.
Private Function AddTreemapCanvas(ws As Worksheet, ByVal baseName As String, ByVal baseLeft As Double, ByVal baseTop As Double) As Shape
    Dim shp As Shape
    Set shp = ws.Shapes.AddShape( _
                    Type:=msoShapeRectangle, _
                    Left:=baseLeft, Top:=baseTop, _
                    Width:=chartWidth, Height:=chartHeight)

    With shp
        .name = baseName & "_Canvas"
        .Fill.Visible = msoTrue
        .Fill.Solid
        .Fill.ForeColor.RGB = colorWhite
        .Line.Visible = msoFalse
        ' The ChartObject already exists (created before chrome), so a just-added
        ' shape sits in front of it. Send the canvas to the back so the chart tiles
        ' and chrome render on top of the white background, not behind it.
        .ZOrder msoSendToBack
    End With

    Set AddTreemapCanvas = shp
End Function


Private Function AddTreemapFigureBox(ws As Worksheet, ByVal baseName As String, ByVal baseLeft As Double, ByVal baseTop As Double) As Shape
    Dim shp As Shape
    Set shp = ws.Shapes.AddTextbox( _
                    Orientation:=msoTextOrientationHorizontal, _
                    Left:=baseLeft, Top:=baseTop + figureBoxTop, _
                    Width:=titleBoxWidth, Height:=figureBoxHeight)

    With shp
        .name = baseName & "_FigureBox"
        .Fill.Visible = msoFalse
        .Line.Visible = msoFalse
        .TextFrame2.TextRange.Text = figureBoxDefaultText
        With .TextFrame2.TextRange.Font
            .Size = figureFontSize
            .name = fontPrimary
            .Fill.ForeColor.RGB = figureFontColor
            .Bold = msoFalse
        End With
    End With

    Set AddTreemapFigureBox = shp
End Function


Private Function AddTreemapTitleBox(ws As Worksheet, ByVal baseName As String, ByVal baseLeft As Double, ByVal baseTop As Double) As Shape
    Dim shp As Shape
    Set shp = ws.Shapes.AddTextbox( _
                    Orientation:=msoTextOrientationHorizontal, _
                    Left:=baseLeft, Top:=baseTop + titleBoxTop, _
                    Width:=titleBoxWidth, Height:=titleBoxHeight)

    With shp
        .name = baseName & "_TitleBox"
        .Fill.Visible = msoFalse
        .Line.Visible = msoFalse
        .TextFrame2.VerticalAnchor = msoAnchorMiddle
        .TextFrame2.TextRange.Text = titleDefaultText
        With .TextFrame2.TextRange.Font
            .Size = titleFontSize
            .name = fontPrimary
            .Fill.ForeColor.RGB = titleFontColor
            .Bold = msoTrue
        End With
    End With

    Set AddTreemapTitleBox = shp
End Function


Private Function AddTreemapSubtitleBox(ws As Worksheet, ByVal baseName As String, ByVal baseLeft As Double, ByVal baseTop As Double) As Shape
    Dim shp As Shape
    Set shp = ws.Shapes.AddTextbox( _
                    Orientation:=msoTextOrientationHorizontal, _
                    Left:=baseLeft, Top:=baseTop + subtitleBoxTop, _
                    Width:=titleBoxWidth, Height:=subtitleBoxHeight)

    With shp
        .name = baseName & "_SubTitleBox"
        .Fill.Visible = msoFalse
        .Line.Visible = msoFalse
        .TextFrame2.VerticalAnchor = msoAnchorMiddle
        .TextFrame2.TextRange.Text = subtitleDefaultText
        With .TextFrame2.TextRange.Font
            .Size = subTitleFontSize
            .Fill.ForeColor.RGB = subTitleFontColor
            .name = fontPrimary
            .Bold = msoFalse
        End With
    End With

    Set AddTreemapSubtitleBox = shp
End Function


Private Function AddTreemapSourceBox(ws As Worksheet, ByVal baseName As String, ByVal baseLeft As Double, ByVal baseTop As Double) As Shape
    ' Sit the source box INSIDE the canvas: its bottom aligns with the canvas bottom
    ' edge (baseTop + chartHeight - sourceBoxHeight) and its left with the canvas left.
    ' The source box and the bottom-right logo share the bottom band: they clear each
    ' other only because sourceBoxWidth (0.8 * canvas) ends left of the logo column.
    ' Keep sourceBoxWidthProportion well below ~0.88 (or shrink the logo) or the source
    ' text will slide under the logo.
    Dim shp As Shape
    Set shp = ws.Shapes.AddTextbox( _
                    msoTextOrientationHorizontal, _
                    baseLeft, baseTop + chartHeight - sourceBoxHeight, sourceBoxWidth, sourceBoxHeight)

    With shp
        .name = baseName & "_SourceBox"
        .Fill.Visible = msoFalse
        .Line.Visible = msoFalse
        .TextFrame.Characters.Text = sourceDefaultText & vbNewLine & notesDefaultText
        .TextFrame.Characters.Font.Size = sourceTextFontSize
        .TextFrame.Characters.Font.name = fontPrimary
        .TextFrame.VerticalAlignment = xlVAlignBottom
        .IncrementLeft -sourceBoxLeftNudge
    End With

    Set AddTreemapSourceBox = shp
End Function


' Returns Nothing (and shows MsgLogoDecodeFailed) if the embedded logo can't be
' decoded - the rest of the chrome is still built.
Private Function AddTreemapLogo(ws As Worksheet, ByVal baseName As String, ByVal baseLeft As Double, ByVal baseTop As Double) As Shape
    On Error GoTo Fail

    Dim tmpPath As String
    tmpPath = Environ$("TEMP") & "\logo_temp.svg"

    If Not Base64ToFile(LogoPNG_Base64, tmpPath) Then
        MsgLogoDecodeFailed
        Set AddTreemapLogo = Nothing
        Exit Function
    End If

    Dim logoShape As Shape
    Set logoShape = ws.Shapes.AddPicture( _
                Filename:=tmpPath, _
                LinkToFile:=msoFalse, _
                SaveWithDocument:=msoTrue, _
                Left:=baseLeft, Top:=baseTop, _
                Width:=-1, Height:=-1)

    logoShape.name = baseName & "_LogoImage"

    ' Scale against the fixed canvas (chart is forced to chartWidth x chartHeight).
    Dim TargetHeight As Single, TargetWidth As Single
    TargetHeight = chartHeight * logoHeightScale
    TargetWidth = TargetHeight * logoAspectRatio

    logoShape.LockAspectRatio = msoFalse
    logoShape.Height = TargetHeight
    logoShape.Width = TargetWidth

    ' Position bottom-right of the chart, in worksheet coordinates.
    logoShape.Left = baseLeft + chartWidth - logoShape.Width - logoMarginRight
    logoShape.Top = baseTop + chartHeight - logoShape.Height - logoMarginBottom

    On Error Resume Next
    Kill tmpPath
    On Error GoTo 0

    Set AddTreemapLogo = logoShape
    Exit Function

Fail:
    MsgLogoDecodeFailed
    Set AddTreemapLogo = Nothing
End Function


' ============================================================
'   GROUPING & LIFECYCLE
' ============================================================

' Groups the chrome shapes + the ChartObject into one named group. On failure the
' shapes are left in place (ungrouped) and the user is told.
Private Function GroupTreemapChrome(ws As Worksheet, cht As Chart, ByRef chromeNames() As String) As Shape
    On Error GoTo Fail

    ' Build the member-name list: the ChartObject's own shape + each chrome shape
    ' that was actually created (logo may be absent). Upper bound 7 = ChartObject +
    ' canvas + 4 text boxes + logo. Typed as Variant because Shapes.Range expects its
    ' index packaged in a Variant - a typed String() array can raise type-mismatch
    ' (error 13) on some Excel builds.
    Dim names() As Variant
    Dim n As Long
    ReDim names(1 To 7)

    n = n + 1: names(n) = cht.Parent.name

    Dim i As Long
    For i = LBound(chromeNames) To UBound(chromeNames)
        If Len(chromeNames(i)) > 0 Then
            n = n + 1: names(n) = chromeNames(i)
        End If
    Next i

    ReDim Preserve names(1 To n)

    Dim grp As Shape
    Set grp = ws.Shapes.Range(names).Group
    grp.name = TreemapGroupName(cht)

    Set GroupTreemapChrome = grp
    Exit Function

Fail:
    MsgTreemapGroupFailed
    Set GroupTreemapChrome = Nothing
End Function


' Removes a prior treemap group and its chrome members so a re-run doesn't duplicate.
' Ungrouping leaves the ChartObject intact (the freshly-retyped chart reference stays
' valid); only the chrome members are deleted.
Private Sub RemoveExistingTreemapChrome(cht As Chart)
    On Error Resume Next

    Dim ws As Worksheet
    Set ws = HostSheet(cht)
    If ws Is Nothing Then Exit Sub

    Dim baseName As String
    baseName = cht.Parent.name

    ' Ungroup the prior group if present (ChartObject survives the ungroup).
    Dim grp As Shape
    Set grp = ws.Shapes(TreemapGroupName(cht))
    If Not grp Is Nothing Then
        If grp.Type = msoGroup Then grp.Ungroup
    End If
    Set grp = Nothing

    ' Delete the prefixed chrome members by name.
    SafeDeleteSheetShape ws, baseName & "_Canvas"
    SafeDeleteSheetShape ws, baseName & "_FigureBox"
    SafeDeleteSheetShape ws, baseName & "_TitleBox"
    SafeDeleteSheetShape ws, baseName & "_SubTitleBox"
    SafeDeleteSheetShape ws, baseName & "_SourceBox"
    SafeDeleteSheetShape ws, baseName & "_LogoImage"

    On Error GoTo 0
End Sub


' ============================================================
'   LAYOUT (public - called by the builder before chrome)
' ============================================================

' Resolves the canvas top-left (the origin all chrome and the chart band are laid
' out from). On a re-run an existing "<chart>_Canvas" shape exists, so reuse its
' position - this keeps the layout stable even if the group was moved. On first run
' there is no canvas yet, so fall back to the chart's current position.
Public Sub TreemapCanvasOrigin(cht As Chart, ByRef outLeft As Double, ByRef outTop As Double)
    outLeft = cht.Parent.Left
    outTop = cht.Parent.Top

    Dim ws As Worksheet
    Set ws = HostSheet(cht)
    If ws Is Nothing Then Exit Sub

    Dim canvas As Shape
    On Error Resume Next
    Set canvas = ws.Shapes(cht.Parent.name & "_Canvas")
    On Error GoTo 0
    If Not canvas Is Nothing Then
        outLeft = canvas.Left
        outTop = canvas.Top
    End If
End Sub


' Positions the treemap ChartObject as a band inside the canvas, leaving the title
' block above and the logo/source below. Mirrors the classic no-legend plot-area
' geometry (modConfig): inset left/right by plotAreaLeft, top at plotAreaTop_noLegend,
' height plotAreaHeight_noLegend - all relative to the canvas origin.
Public Sub PositionTreemapChart(cht As Chart, ByVal baseLeft As Double, ByVal baseTop As Double)
    With cht.Parent
        .Left = baseLeft + plotAreaLeft
        .Top = baseTop + plotAreaTop_noLegend
        .Width = chartWidth - 2 * plotAreaLeft
        .Height = plotAreaHeight_noLegend
    End With
End Sub


' ============================================================
'   HELPERS
' ============================================================

' The host worksheet for an embedded chart, or Nothing for a chart sheet
' (cht.Parent is the Workbook) - chart sheets are out of scope.
Private Function HostSheet(cht As Chart) As Worksheet
    On Error Resume Next
    If TypeName(cht.Parent) = "ChartObject" Then
        Set HostSheet = cht.Parent.Parent
    End If
    On Error GoTo 0
End Function


' Deterministic group name for a given chart.
' Precondition: cht.Parent is a ChartObject (embedded chart). All callers reach this
' only after BuildTreemapChrome's HostSheet guard has passed.
Private Function TreemapGroupName(cht As Chart) As String
    TreemapGroupName = treemapGroupPrefix & cht.Parent.name
End Function


' Deletes a worksheet shape by name if present (mirrors SafeDeleteShape, which only
' searches cht.Shapes).
Private Sub SafeDeleteSheetShape(ws As Worksheet, ByVal nm As String)
    On Error Resume Next
    ws.Shapes(nm).Delete
    On Error GoTo 0
End Sub
