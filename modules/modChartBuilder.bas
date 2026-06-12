Attribute VB_Name = "modChartBuilder"
Option Explicit

' Shared formatting pipeline applied to every chart type.
' colorMode: "FILL" for bar/column charts; "LINE" for line/slope/scatter charts.
' defaults: ChartDefaults UDT containing all five formatting options.
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
Public Sub ApplyChartPipeline(cht As Chart, ByVal colorMode As String, ByRef defaults As ChartDefaults)
    Call OuterFormat(cht, defaults)
    Call FormatXAxisTitle(cht)
    Call InsertLogo(cht)
    Call InsertSource(cht)
    Call FormatTitle(cht)
    Call FormatGridlines(cht)
    Call FormatXAxis(cht)
    Call FormatSeriesColors(cht, UCase$(colorMode))
    Call ApplyDefaultFormatting(cht, defaults)
End Sub


Function OuterFormat(cht As Chart, ByRef defaults As ChartDefaults) As Boolean
    On Error GoTo Fail

    Dim SeriesCount As Long

    'Font
    cht.ChartArea.Font.name = fontPrimary

    'Format Y-axis line to white using Axis.Border (no Select required)
    If cht.HasAxis(xlValue, xlPrimary) Then
        With cht.Axes(xlValue, xlPrimary).Border
            .LineStyle = xlContinuous
            .Color = colorWhite
            .Weight = axisLineWeight
        End With
    End If

    'Format X-axis line to white using Axis.Border (no Select required)
    If cht.HasAxis(xlCategory, xlPrimary) Then
        FormatAxisLineWhite cht.Axes(xlCategory, xlPrimary)
    End If

    'Remove axis titles
    RemoveAxisTitles cht

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
    SeriesCount = cht.SeriesCollection.Count
    If SeriesCount = 0 Then
        OuterFormat = True
        Exit Function
    End If

    'Plot area adjustments
    ApplyPlotAreaGeometry cht, SeriesCount, defaults.Legend

    OuterFormat = True
    Exit Function

Fail:
    OuterFormat = False
End Function


Private Sub FormatAxisLineWhite(ax As Axis)
    'Format axis line to white using Axis.Border (no Select required).
    On Error Resume Next
    With ax.Border
        .LineStyle = xlContinuous
        .Color = colorWhite
        .Weight = axisLineWeight
    End With
    On Error GoTo 0
End Sub


Private Sub RemoveAxisTitles(cht As Chart)
    If cht.HasAxis(xlValue) Then
        If cht.Axes(xlValue).HasTitle Then cht.Axes(xlValue).AxisTitle.Delete
    End If
    If cht.HasAxis(xlCategory) Then
        If cht.Axes(xlCategory).HasTitle Then cht.Axes(xlCategory).AxisTitle.Delete
    End If
End Sub


Private Sub ApplyPlotAreaGeometry(cht As Chart, ByVal SeriesCount As Long, ByVal ShowLegend As Boolean)
    Dim pa As PlotArea
    Set pa = cht.PlotArea

    Dim HasMultipleSeries As Boolean, HasLegend As Boolean
    HasMultipleSeries = (SeriesCount > 1)
    HasLegend = HasMultipleSeries And ShowLegend And cht.hasLegend

    'Remove legend if single series or legend disabled in defaults
    If Not HasLegend And cht.hasLegend Then
        cht.Legend.Delete
    End If

    'Position legend and adjust plot area
    If HasLegend Then
        cht.Legend.Position = xlLegendPositionTop
        cht.Legend.Left = legendLeftPad
        cht.Legend.Font.Color = legendFontColor

        With pa
            .Height = plotAreaHeight
            .Top = plotAreaTop
            .Width = plotAreaWidth
            .Left = plotAreaLeft
        End With
    Else
        With pa
            .Height = plotAreaHeight_noLegend
            .Top = plotAreaTop_noLegend
            .Width = plotAreaWidth
            .Left = plotAreaLeft
        End With
    End If
End Sub


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

            .Top = legendTop
            .Left = legendLeftPad
        End With
    Else
        ' No legend — use no-legend plot area dimensions
        With plt
            .Height = plotAreaHeight_noLegend
            .Top = plotAreaTop_noLegend
        End With
    End If

    FormatXAxisTitle = True
    Exit Function

Fail:
    FormatXAxisTitle = False
End Function


Public Function InsertLogo(cht As Chart) As Boolean
    On Error GoTo Fail

    'Decode Base64 to temp file
    Dim tmpPath As String
    tmpPath = Environ$("TEMP") & "\logo_temp.svg"

    If Not Base64ToFile(LogoPNG_Base64, tmpPath) Then
        MsgLogoDecodeFailed
        InsertLogo = False
        Exit Function
    End If

    'Remove existing logo to avoid duplicates
    SafeDeleteShape cht, "LogoImage"

    'Insert logo at native size
    Dim logoShape As Shape
    Set logoShape = cht.Shapes.AddPicture( _
                Filename:=tmpPath, _
                LinkToFile:=msoFalse, _
                SaveWithDocument:=msoTrue, _
                Left:=0, Top:=0, _
                Width:=-1, Height:=-1)

    logoShape.name = "LogoImage"

    'Scale to target dimensions
    Dim ChartWidth As Single, ChartHeight As Single
    ChartWidth = cht.Parent.Width
    ChartHeight = cht.Parent.Height

    Dim TargetHeight As Single, TargetWidth As Single
    TargetHeight = ChartHeight * logoHeightScale
    TargetWidth = TargetHeight * logoAspectRatio

    logoShape.LockAspectRatio = msoFalse
    logoShape.Height = TargetHeight
    logoShape.Width = TargetWidth

    'Position bottom right
    logoShape.Left = ChartWidth - logoShape.Width - logoMarginRight
    logoShape.Top = ChartHeight - logoShape.Height - logoMarginBottom

    'Clean up temp file
    On Error Resume Next
    Kill tmpPath
    On Error GoTo 0

    InsertLogo = True
    Exit Function

Fail:
    InsertLogo = False
    MsgError "InsertLogo"
End Function


Function InsertSource(cht As Chart) As Boolean
    On Error GoTo Fail

    SafeDeleteShape cht, "SourceBox"

    'Add textbox at bottom-left
    Dim SourceBox As Shape
    Dim ChartHeight As Long
    ChartHeight = cht.Parent.Height

    Set SourceBox = cht.Shapes.AddTextbox( _
                    msoTextOrientationHorizontal, _
                    0, ChartHeight, sourceBoxWidth, sourceBoxHeight)

    With SourceBox
        .name = "SourceBox"
        .TextFrame.Characters.Text = sourceDefaultText & vbNewLine & notesDefaultText
        .TextFrame.Characters.Font.Size = sourceTextFontSize
        .TextFrame.Characters.Font.name = fontPrimary
        .TextFrame.VerticalAlignment = xlVAlignBottom
        .IncrementLeft -sourceBoxLeftNudge
    End With

    InsertSource = True
    Exit Function

Fail:
    InsertSource = False
    MsgError "InsertSource"
End Function



Function FormatTitle(cht As Chart) As Boolean
    On Error GoTo Fail

    'Delete existing title-related boxes
    SafeDeleteShape cht, "FigureBox"
    SafeDeleteShape cht, "TitleBox"
    SafeDeleteShape cht, "SubTitleBox"
    SafeDeleteShape cht, "YAxisLabelBox"

    'Remove built-in chart title
    If cht.HasTitle Then cht.ChartTitle.Delete

    'Create all title boxes
    CreateFigureBox cht
    CreateTitleBox cht
    CreateSubtitleBox cht
    CreateYAxisLabelBox cht, cht.hasLegend

    FormatTitle = True
    Exit Function

Fail:
    FormatTitle = False
End Function


Private Sub CreateFigureBox(cht As Chart)
    Dim shp As Shape
    Set shp = cht.Shapes.AddTextbox( _
                    Orientation:=msoTextOrientationHorizontal, _
                    Left:=0, Top:=figureBoxTop, Width:=titleBoxWidth, Height:=figureBoxHeight)

    With shp
        .name = "FigureBox"
        .TextFrame2.TextRange.Text = figureBoxDefaultText
        With .TextFrame2.TextRange.Font
            .Size = figureFontSize
            .name = fontPrimary
            .Fill.ForeColor.RGB = figureFontColor
            .Bold = msoFalse
        End With
        .Top = .Top - titleBoxNudge
        .Left = .Left - titleBoxNudge
    End With
End Sub


Private Sub CreateTitleBox(cht As Chart)
    Dim shp As Shape
    Set shp = cht.Shapes.AddTextbox( _
                    Orientation:=msoTextOrientationHorizontal, _
                    Left:=0, Top:=titleBoxTop, Width:=titleBoxWidth, Height:=titleBoxHeight)

    With shp
        .name = "TitleBox"
        .TextFrame2.VerticalAnchor = msoAnchorMiddle
        .TextFrame2.TextRange.Text = titleDefaultText
        With .TextFrame2.TextRange.Font
            .Size = titleFontSize
            .name = fontPrimary
            .Fill.ForeColor.RGB = titleFontColor
            .Bold = msoTrue
        End With
        .Top = .Top - titleBoxNudge
        .Left = .Left - titleBoxNudge
    End With
End Sub


Private Sub CreateSubtitleBox(cht As Chart)
    Dim shp As Shape
    Set shp = cht.Shapes.AddTextbox( _
                    Orientation:=msoTextOrientationHorizontal, _
                    Left:=0, Top:=subtitleBoxTop, Width:=titleBoxWidth, Height:=subtitleBoxHeight)

    With shp
        .name = "SubTitleBox"
        .TextFrame2.TextRange.Text = subtitleDefaultText
        With .TextFrame2.TextRange.Font
            .Size = subTitleFontSize
            .Fill.ForeColor.RGB = subTitleFontColor
            .name = fontPrimary
            .Bold = msoFalse
        End With
        .Top = .Top - titleBoxNudge
        .Left = .Left - titleBoxNudge
    End With
End Sub


Private Sub CreateYAxisLabelBox(cht As Chart, ByVal HasLegend As Boolean)
    Dim shp As Shape
    Dim yAxisTop As Single

    yAxisTop = IIf(HasLegend, yAxisLabelTop, yAxisLabelTop_noLegend)

    Set shp = cht.Shapes.AddTextbox( _
                    Orientation:=msoTextOrientationHorizontal, _
                    Left:=0, Top:=yAxisTop, Width:=titleBoxWidth, Height:=yAxisLabelHeight)

    With shp
        .name = "YAxisLabelBox"
        .TextFrame2.TextRange.Text = yAxisDefaultText
        With .TextFrame2.TextRange.Font
            .Size = axisFontSize
            .name = fontPrimaryItalic
            .Bold = msoFalse
            .Italic = msoTrue
        End With
        .Left = .Left - titleBoxNudge
    End With
End Sub


Function FormatGridlines(cht As Chart) As Boolean
    On Error GoTo Fail

    If Not cht.HasAxis(xlValue) Then
        FormatGridlines = True
        Exit Function
    End If

    Dim ax As Axis
    Set ax = cht.Axes(xlValue)

    ' Add gridlines if missing
    If Not ax.HasMajorGridlines Then
        cht.SetElement msoElementPrimaryValueGridLinesMajor
    End If

    ' Apply major gridline formatting
    With ax.MajorGridlines.Format.Line
        .Visible = msoTrue
        .Weight = gridlineWeight
        .DashStyle = msoLineSolid
        .ForeColor.RGB = colorNeutral2
    End With

    FormatGridlines = True
    Exit Function

Fail:
    FormatGridlines = False
End Function


Function FormatXAxis(cht As Chart) As Boolean
    On Error GoTo Fail

    'Format category (X) axis
    If cht.HasAxis(xlCategory) Then
        With cht.Axes(xlCategory)
            .TickLabels.Font.Size = axisFontSize
            .TickLabels.Font.Color = legendFontColor
        End With
        FormatCategoryAxisLine cht.Axes(xlCategory)
    End If

    'Format value (Y) axis
    If cht.HasAxis(xlValue) Then
        With cht.Axes(xlValue)
            .TickLabels.Font.Size = axisFontSize
            .TickLabels.Font.Color = axisFontColor
        End With
    End If

    FormatXAxis = True
    Exit Function

Fail:
    FormatXAxis = False
End Function


Private Sub FormatCategoryAxisLine(ax As Axis)
    'Format category axis line to white using Axis.Border object (no Select required).
    On Error Resume Next
    With ax.Border
        .LineStyle = xlContinuous
        .Color = colorWhite
        .Weight = axisLineWeight
    End With
    On Error GoTo 0
End Sub


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


' Applies defaults after the full pipeline completes.
' Undoes gridlines added by FormatGridlines and removes axes per defaults.AxisDisplay.
' Called as the final step of ApplyChartPipeline so earlier steps can still access axes.
Private Sub ApplyDefaultFormatting(cht As Chart, ByRef defaults As ChartDefaults)
    On Error GoTo CleanFail

    ' --- Gridlines: remove value-axis gridlines unless Y or Both are requested
    If defaults.Gridlines = axisNone Or defaults.Gridlines = axisX Then
        If cht.HasAxis(xlValue) Then
            If cht.Axes(xlValue).HasMajorGridlines Then
                cht.Axes(xlValue).MajorGridlines.Delete
            End If
        End If
    End If

    ' --- Axis display: show/hide each axis per defaults parameter
    Dim showX As Boolean, showY As Boolean
    showX = (defaults.AxisDisplay = axisX Or defaults.AxisDisplay = axisBoth)
    showY = (defaults.AxisDisplay = axisY Or defaults.AxisDisplay = axisBoth)

    cht.HasAxis(xlValue) = showY
    cht.HasAxis(xlCategory) = showX

    Exit Sub
CleanFail:
    MsgError "ApplyDefaultFormatting"
End Sub


'Returns the chart to style, modifying it in-place. Two entry paths:
'  1. A chart is already active  → retype it to chartType; return it.
'  2. A range is selected        → create a new chart of chartType; return it.
'Returns Nothing on any other selection state or error.
Public Function GetTargetChart(ByVal chartType As Long) As Chart
    On Error GoTo Fail

    Set GetTargetChart = Nothing

    'Path 1: Retype active chart
    If Not ActiveChart Is Nothing Then
        ActiveChart.chartType = chartType
        Set GetTargetChart = ActiveChart
        Exit Function
    End If

    'Path 2: Create new chart from selected range
    If TypeName(Selection) <> "Range" Then
        MsgSelectRangeOrChart
        Exit Function
    End If

    On Error Resume Next
    ActiveSheet.Shapes.AddChart2(-1, chartType).Select
    If Not ActiveChart Is Nothing Then
        Set GetTargetChart = ActiveChart
    End If
    On Error GoTo 0

    Exit Function

Fail:
    'Already set to Nothing on entry
End Function
