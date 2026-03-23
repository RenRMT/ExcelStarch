Attribute VB_Name = "modChartBuilder"
Option Explicit

' Shared formatting pipeline applied to every chart type.
' colorMode: "FILL" for bar/column charts; "LINE" for line/slope/scatter charts.
'
' Step order matters:
'   1. OuterFormat    — sets chart size and plot area geometry first; everything else depends on it
'   2. FormatXAxisTitle — positions relative to plot area InsideTop/InsideHeight, so must follow OuterFormat
'   3. InsertLogo     — anchored to chart bottom-right; independent of plot area
'   4. InsertSource   — anchored to chart bottom-left; must exist before FormatTitle so boxes don't overlap
'   5. FormatTitle    — adds title/subtitle/y-axis label text boxes at top-left
'   6. FormatGridlines — applies major gridline style to value axis
'   7. FormatXAxis    — sizes and colors axis tick labels; runs after gridlines to avoid selection conflicts
'   8. FormatSeriesColors — applied last so series exist and pipeline hasn't altered their format
'
' Chart types that skip steps (slope, dot plot, scatter) call individual functions directly.
Public Sub ApplyChartPipeline(cht As Chart, ByVal colorMode As String)
    Call OuterFormat(cht)
    Call FormatXAxisTitle(cht)
    Call InsertLogo(cht)
    Call InsertSource(cht)
    Call FormatTitle(cht)
    Call FormatGridlines(cht)
    Call FormatXAxis(cht)
    Call FormatSeriesColors(cht, UCase$(colorMode))
End Sub


Function OuterFormat(cht As Chart) As Boolean
    On Error GoTo Fail

    Dim seriescount As Long

    'Font
    cht.ChartArea.Font.name = fontPrimary

    'Hide Y-axis line
    If cht.HasAxis(xlValue) Then
        cht.Axes(xlValue).Format.Line.Visible = msoFalse
    End If

    'Hide X-axis line (requires Select — Excel doesn't expose Format.Line on Axis directly)
    If cht.HasAxis(xlCategory) Then
        cht.Axes(xlCategory).Select
        Selection.Format.Line.Visible = msoFalse
    End If

    'Remove axis titles
    If cht.HasAxis(xlValue) Then
        If cht.Axes(xlValue).HasTitle Then cht.Axes(xlValue).AxisTitle.Delete
    End If
    If cht.HasAxis(xlCategory) Then
        If cht.Axes(xlCategory).HasTitle Then cht.Axes(xlCategory).AxisTitle.Delete
    End If

    'Chart size
    If TypeName(cht.Parent) = "ChartObject" Then
        With cht.Parent
            .Width = chartWidth
            .Height = chartHeight
        End With
    End If

    'Remove border
    cht.ChartArea.Border.LineStyle = xlNone

    'Series count
    If cht.SeriesCollection.Count = 0 Then Exit Function
    seriescount = cht.SeriesCollection.Count

    ' Plot area adjustments
    Dim pa As PlotArea
    Set pa = cht.PlotArea

    If seriescount = 1 Then

        'Remove legend
        If cht.hasLegend Then cht.Legend.Delete

        pa.Height = plotAreaHeight
        pa.Top = plotAreaTop_default
        pa.Width = plotAreaWidth
        pa.Left = plotAreaLeft

    Else

        If cht.hasLegend Then

            cht.Legend.Position = xlLegendPositionTop
            cht.Legend.Left = legend_leftPad
            cht.Legend.Font.color = legendFontColor

            pa.Height = plotAreaHeight
            pa.Top = plotAreaTop_default
            pa.Width = plotAreaWidth
            pa.Left = plotAreaLeft

        Else

            pa.Height = plotAreaHeight
            pa.Top = plotAreaTop_noLegend
            pa.Width = plotAreaWidth
            pa.Left = plotAreaLeft

        End If

    End If

    OuterFormat = True
    Exit Function

Fail:
    OuterFormat = False
End Function


Function FormatXAxisTitle(cht As Chart) As Boolean
    On Error GoTo Fail

    Dim shp As Shape
    Dim plt As PlotArea
    Dim tr As TextRange2
    Dim seriescount As Long

    Set plt = cht.PlotArea
    seriescount = cht.SeriesCollection.Count

    ' Remove existing XAxisBox if present
    SafeDeleteShape cht, "XAxisBox"

    ' Create X-axis title textbox
    Set shp = cht.Shapes.AddTextbox( _
                Orientation:=msoTextOrientationHorizontal, _
                Left:=10, Top:=10, Width:=100, Height:=2)

    shp.name = "XAxisBox"

    Set tr = shp.TextFrame2.TextRange
    tr.Text = xAxisDefaultText

    With tr.Font
        .Italic = msoTrue
        .Size = axisFontSize
        .Fill.ForeColor.RGB = axisFontColor
        .name = fontPrimary
    End With

    With shp.TextFrame2
        .VerticalAnchor = msoAnchorMiddle
        .WordWrap = msoFalse
        .AutoSize = msoAutoSizeShapeToFitText
    End With

    ' Position below plot area, centered.
    ' InsideTop/InsideHeight refer to the inner plot boundary (excluding axis tick labels),
    ' so this places the title just below where the data ends, not below the axis labels.
    shp.Top = plt.InsideTop + plt.InsideHeight + xAxisTitle_plotGap
    shp.Left = plt.InsideLeft + (plt.InsideWidth - shp.Width) / 2

    ' Legend repositioning
    If cht.hasLegend Then
        With cht.Legend
            .Font.Size = axisFontSize

            .Top = legend_top
            .Left = legend_leftPad
        End With
    Else
        ' Adjust plot area when no legend
        With plt
            If seriescount = 1 Then
                .Height = plotArea_noLegendSingleHeight
                .Top = plotArea_noLegendSingleTop
            Else
                .Height = plotArea_noLegendMultiHeight
                .Top = plotArea_noLegendMultiTop
            End If
        End With
    End If

    FormatXAxisTitle = True
    Exit Function

Fail:
    FormatXAxisTitle = False
End Function


Public Function InsertLogo(cht As Chart) As Boolean
    On Error GoTo Fail

    ' 1. Decode Base64 -> temp file
    Dim tmp As String
    tmp = Environ$("TEMP") & "\logo_temp.svg"

    If Not Base64ToFile(LogoPNG_Base64, tmp) Then
        MsgLogoDecodeFailed
        Exit Function
    End If

    ' 2. Remove existing logo (avoid duplicates)
    Dim s As Shape
    For Each s In cht.Shapes
        If s.name = "LogoImage" Then s.Delete
    Next s

    ' 3. Insert at native size, no locking
    Dim shp As Shape
    Set shp = cht.Shapes.AddPicture( _
                Filename:=tmp, _
                LinkToFile:=msoFalse, _
                SaveWithDocument:=msoTrue, _
                Left:=0, Top:=0, _
                Width:=-1, Height:=-1)

    shp.name = "LogoImage"

    ' 4. Apply target dimensions: height = logoHeightScale x chart height, width = logoAspectRatio x height
    Dim chW As Single, chH As Single
    chW = cht.Parent.Width
    chH = cht.Parent.Height

    Dim targetH As Single, targetW As Single
    targetH = chH * logoHeightScale
    targetW = targetH * logoAspectRatio

    ' Must unlock aspect ratio so we can apply our own ratio
    shp.LockAspectRatio = msoFalse

    shp.Height = targetH
    shp.Width = targetW

    ' 5. Position bottom right
    shp.Left = chW - shp.Width - logoMarginRight
    shp.Top = chH - shp.Height - logoMarginBottom

    On Error Resume Next
    Kill tmp
    On Error GoTo 0

    InsertLogo = True
    Exit Function

Fail:
    MsgError "InsertLogo"
End Function


Function InsertSource(cht As Chart) As Boolean
    On Error GoTo Fail

    Dim sourceB As TextBox
    Dim chHeight As Long

    SafeDeleteShape cht, "SourceBox"

    'Chart dimensions
    chHeight = cht.Parent.Height

    'Add textbox at bottom-left
    Set sourceB = cht.TextBoxes.Add(0, chHeight, sourceBoxWidth, sourceBoxHeight)

    With sourceB
        .name = "SourceBox"
        .Text = sourceDefaultText & vbNewLine & notesDefaultText
        .Font.Size = sourceTextFontSize
        .Font.name = fontPrimary
    End With

    'Bottom-align the text
    cht.Shapes.Range(Array("SourceBox")).Select
    With Selection
        .VerticalAlignment = xlBottom
    End With

    ' padding
    cht.Shapes.Range(Array("SourceBox")).Select
    Selection.ShapeRange.IncrementLeft -sourceBoxLeftNudge

    InsertSource = True
    Exit Function

Fail:
    MsgError "InsertSource"
End Function


Function FormatTitle(cht As Chart) As Boolean
    On Error GoTo Fail

    Dim titleB1 As Shape, titleB2 As Shape, titleB3 As Shape
    Dim plt As PlotArea
    Dim seriescount As Long
    Dim yAxisTop As Single
    Dim hasLegend As Boolean

    ' Delete existing title-related boxes safely
    SafeDeleteShape cht, "TitleBox"
    SafeDeleteShape cht, "SubTitleBox"
    SafeDeleteShape cht, "YAxisLabelBox"

    ' Remove built-in chart title if present
    If cht.HasTitle Then cht.ChartTitle.Delete

    ' Capture chart state
    seriescount = cht.SeriesCollection.Count
    hasLegend = cht.hasLegend
    Set plt = cht.PlotArea

    ' Title
    Set titleB1 = cht.Shapes.AddTextbox( _
                    Orientation:=msoTextOrientationHorizontal, _
                    Left:=0, Top:=0, Width:=titleBoxWidth, Height:=titleBoxHeight)

    With titleB1
        .name = "TitleBox"
        With .TextFrame2.TextRange
            .Text = titleDefaultText
            With .Font
                .Size = titleFontSize
                .name = fontPrimary
                .Fill.ForeColor.RGB = colorBrand1
                .Bold = msoTrue
            End With
        End With

        ' nudge
        .Top = .Top - titleBoxNudge
        .Left = .Left - titleBoxNudge
    End With

    ' Subtitle
    Set titleB2 = cht.Shapes.AddTextbox( _
                    Orientation:=msoTextOrientationHorizontal, _
                    Left:=0, Top:=subtitleBoxTop, Width:=titleBoxWidth, Height:=subtitleBoxHeight)

    With titleB2
        .name = "SubTitleBox"
        With .TextFrame2.TextRange
            .Text = subtitleDefaultText
            With .Font
                .Size = subTitleFontSize
                .Fill.ForeColor.RGB = colorBrand2
                .name = fontPrimary
                .Bold = msoFalse
            End With
        End With

        ' nudge
        .Top = .Top - titleBoxNudge
        .Left = .Left - titleBoxNudge
    End With

    ' Y-axis Label
    If hasLegend Then
        yAxisTop = yAxisLabel_legendTop
    Else
        yAxisTop = IIf(seriescount = 1, yAxisLabel_singleTop, yAxisLabel_multiTop)
    End If

    Set titleB3 = cht.Shapes.AddTextbox( _
                    Orientation:=msoTextOrientationHorizontal, _
                    Left:=0, Top:=yAxisTop, Width:=titleBoxWidth, _
                    Height:=IIf(hasLegend, yAxisLabel_legendHeight, yAxisLabel_noLegendHeight))

    With titleB3
        .name = "YAxisLabelBox"
        With .TextFrame2.TextRange
            .Text = yAxisDefaultText
            With .Font
                .Size = axisFontSize
                .name = fontPrimaryItalic
                .Bold = msoFalse
                .Italic = msoTrue
            End With
        End With

        ' Left nudge only (matching original logic)
        .Left = .Left - titleBoxNudge
    End With

    FormatTitle = True
    Exit Function

Fail:
    FormatTitle = False
End Function


Function FormatGridlines(cht As Chart) As Boolean
    On Error GoTo Fail

    Dim ax As Axis
    Set ax = cht.Axes(xlValue)

    ' Add gridlines if missing
    If Not ax.HasMajorGridlines Then
        cht.SetElement msoElementPrimaryValueGridLinesMajor
    End If

    ' Apply major gridline formatting
    With ax.MajorGridlines.Format.Line
        .Visible = msoTrue
        .weight = gridlineWeight
        .DashStyle = msoLineSolid
        .ForeColor.rgb = colorNeutral2
    End With

    FormatGridlines = True
    Exit Function

Fail:
    FormatGridlines = False
End Function


Function FormatXAxis(cht As Chart) As Boolean
    On Error GoTo Fail

    'Format size of x-axis & y-axis tick mark labels
    If cht.HasAxis(xlCategory) = True Then
        cht.Axes(xlCategory).TickLabels.Font.Size = axisFontSize

        'Change color of x-axis and y-axis text to black (affects 2013 & 2016)
        cht.Axes(xlCategory, xlPrimary).TickLabels.Font.color = legendFontColor

        'Change x-axis line color
        ' Note: Excel does not expose Format.Line on an Axis object directly —
        ' the property is only accessible via Selection after .Select
        cht.Axes(xlCategory).Select
        With Selection.Format.Line
            .Visible = msoFalse
            .ForeColor.TintAndShade = 0
            .ForeColor.Brightness = 0
            .weight = axisLineWeight
        End With

    End If

    If cht.HasAxis(xlValue) = True Then
        cht.Axes(xlValue).TickLabels.Font.Size = axisFontSize

        'Change color of x-axis and y-axis text to black (affects 2013 & 2016)
        cht.Axes(xlValue, xlPrimary).TickLabels.Font.color = colorBrand3

    End If

    FormatXAxis = True
    Exit Function

Fail:
    FormatXAxis = False
End Function


Function RemoveShadow(cht As Chart) As Boolean
    On Error GoTo Fail

    Dim i As Long
    Dim seriescount As Long

    ' Get series count safely
    seriescount = cht.SeriesCollection.Count
    If seriescount = 0 Then
        RemoveShadow = True
        Exit Function
    End If

    ' Remove shadow directly
    For i = 1 To seriescount
        With cht.SeriesCollection(i).Format.Shadow
            .Visible = msoFalse
        End With
    Next i

    RemoveShadow = True
    Exit Function

Fail:
    RemoveShadow = False
End Function


Public Sub SafeDeleteShape(cht As Chart, ByVal nm As String)
    On Error Resume Next
    cht.Shapes(nm).Delete
    On Error GoTo 0
End Sub


' Returns a duplicate chart to style. Two entry paths:
'   1. A chart is already active  → duplicate it; return the copy.
'   2. A range is selected        → create a new chart of chartType, duplicate it, return the copy.
' In both paths the original is left untouched. Returns Nothing on any other selection state.
Public Function GetTargetChart(ByVal chartType As Long) As Chart
    On Error GoTo Fail

    If Not ActiveChart Is Nothing Then
        ActiveChart.Parent.Duplicate.Select
        If Not ActiveChart Is Nothing Then Set GetTargetChart = ActiveChart
        Exit Function
    End If

    If TypeName(Selection) <> "Range" Then
        MsgSelectRangeOrChart
        Exit Function
    End If

    ActiveSheet.Shapes.AddChart2(-1, chartType).Select
    ActiveChart.Parent.Duplicate.Select

    If Not ActiveChart Is Nothing Then Set GetTargetChart = ActiveChart
    Exit Function

Fail:
    Set GetTargetChart = Nothing
End Function
