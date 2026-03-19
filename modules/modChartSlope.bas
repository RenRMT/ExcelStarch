Attribute VB_Name = "modChartSlope"
Option Explicit

Private Sub BuildSlopeChart()
    Dim inputValue As Variant
    Dim maxcell As Double
    Dim seriescount As Long
    Dim iseries As Long
    Dim imarker As Long
    Dim cht As Chart

    inputValue = InputBox( _
        "How many groups (i.e., rows) do you want in your slope chart? (Must be at least 2 groups.)" & vbNewLine & vbCrLf & _
        "Styling slope charts will sometimes crash in Excel; if that occurs, simply delete the created chart and run it again." & vbNewLine & vbCrLf & _
        "Note that for slope charts with more than 6 series, those series will be assigned a gray color.", _
        "Input Box Text", "2")

    If inputValue = "" Then
        MsgBox "Slope Chart creation cancelled."
        Exit Sub
    End If
    If inputValue < 2 Then
        MsgBox "You have entered too few groups. Please run again with at least 2 groups.", vbExclamation
        Exit Sub
    End If

    maxcell = inputValue

    ' Create new sheet with sample data
    Sheets.Add After:=ActiveSheet
    ActiveSheet.Name = WorksheetFunction.Text(Now(), "h_mm_ss")

    Range("A2").Select
    ActiveCell.FormulaR1C1 = "Group 1"
    Selection.AutoFill Destination:=Range("A2:A" & maxcell + 1), Type:=xlFillDefault
    Range("A2:A" & maxcell).Select

    Range("B2").Select
    ActiveCell.FormulaR1C1 = "1"
    Range("B3").Select
    ActiveCell.FormulaR1C1 = "=R[-1]C+1"
    Range("B3").Select
    If inputValue > 2 Then
        Selection.AutoFill Destination:=Range("B3:B" & maxcell + 1), Type:=xlFillDefault
    End If
    Range("B1").Select
    ActiveCell.FormulaR1C1 = "Field A"
    ActiveCell.HorizontalAlignment = xlRight

    Range("C2").Select
    ActiveCell.FormulaR1C1 = "3"
    Range("C3").Select
    ActiveCell.FormulaR1C1 = "=R[-1]C+1"
    Range("C3").Select
    If inputValue > 2 Then
        Selection.AutoFill Destination:=Range("C3:C" & maxcell + 1), Type:=xlFillDefault
    End If
    Range("C1").Select
    ActiveCell.FormulaR1C1 = "Field B"
    ActiveCell.HorizontalAlignment = xlRight

    ' Create line chart from sample data
    Range("A1:C" & maxcell + 1).Select
    ActiveSheet.Shapes.AddChart.Select
    ActiveChart.ChartType = xlLine
    ActiveChart.SetSourceData Source:=Range("A1:C" & maxcell + 1)
    ActiveChart.PlotBy = xlRows
    ActiveChart.Axes(xlCategory).AxisBetweenCategories = False
    ActiveChart.Axes(xlValue).MajorGridlines.Select
    Selection.Delete
    ActiveChart.Axes(xlValue).Select
    Selection.Delete

    With ActiveChart
        .Parent.Top = dotPlotChartTop
        .Parent.Left = dotPlotChartLeft
    End With

    ' Select and duplicate the chart
    ActiveSheet.ChartObjects(1).Activate
    ActiveChart.Parent.Duplicate.Select

    Set cht = ActiveChart

    ' Partial pipeline (FormatXAxis skipped — slope applies tick label sizes in SlopeChartStyles)
    OuterFormat cht
    FormatXAxisTitle cht
    InsertLogo cht
    InsertSource cht
    FormatTitle cht
    FormatGridlines cht
    FormatSeriesColors cht, "LINE"

    cht.Axes(xlCategory).AxisBetweenCategories = False
    cht.Axes(xlCategory).MajorTickMark = xlTickMarkOutside
    cht.Axes(xlCategory).MinorTickMark = xlTickMarkNone
    cht.Axes(xlValue).MajorTickMark = xlTickMarkOutside
    cht.Axes(xlValue).MinorTickMark = xlTickMarkNone

    If cht.Axes(xlValue).HasMajorGridlines Then
        cht.Axes(xlValue).MajorGridlines.Delete
    End If

    ' X-axis line color
    With cht.Axes(xlCategory).Format.Line
        .Visible = msoTrue
        .ForeColor.rgb = colorBlack
        .ForeColor.TintAndShade = 0
        .ForeColor.Brightness = 0
        .Weight = axisLineWeight
    End With

    seriescount = cht.SeriesCollection.Count

    ' Left (first) point labels
    For iseries = 1 To seriescount
        With cht.SeriesCollection(iseries).Points(1)
            .ApplyDataLabels
            With .DataLabel
                .Position = xlLabelPositionLeft
                .Font.Size = dotPlotLabelFontSize
                .ShowSeriesName = True
                .Separator = " "
            End With
        End With
    Next iseries

    ' Right (last) point labels
    For iseries = 1 To seriescount
        With cht.SeriesCollection(iseries).Points(2)
            .ApplyDataLabels
            With .DataLabel
                .Position = xlLabelPositionRight
                .Font.Size = dotPlotLabelFontSize
                .ShowSeriesName = False
                .Separator = " "
            End With
        End With
    Next iseries

    ' Markers: circle, white fill
    For imarker = 1 To seriescount
        With cht.SeriesCollection(imarker)
            .MarkerStyle = xlMarkerStyleCircle
            .MarkerSize = scatterMarkerSize
            With .Format.Fill
                .Visible = msoTrue
                .Solid
                .ForeColor.rgb = colorWhite
            End With
        End With
    Next imarker

    SlopeChartStyles cht

    Range("A1").Select
End Sub

Private Sub SlopeChartStyles(cht As Chart)
    Dim seriescount As Long
    Dim iseries As Long

    seriescount = cht.SeriesCollection.Count

    ' Tick label size (FormatXAxis skipped for slope charts)
    cht.Axes(xlCategory).TickLabels.Font.Size = axisFontSize

    ' Format left (first) and right (last) point labels per series.
    ' Label text is formatted as "<SeriesName> <Value>" with a space separator.
    ' Character offsets differ by series count because rank numbers 1-9 occupy 1 digit,
    ' while 10+ occupy 2 — shifting where the value part begins:
    '   iseries  1-9:  bold chars 1-7 (name), secondary chars 8-2 (value)
    '   iseries 10+:   bold chars 1-8 (name), secondary chars 10-2 (value)
    For iseries = 1 To seriescount
        cht.SeriesCollection(iseries).Points(1).DataLabel.Select
        ' DoEvents forces pause; without this, code crashes in Excel 2016 PC
        DoEvents: DoEvents: DoEvents: DoEvents: DoEvents: DoEvents
        If iseries < 10 Then
            Selection.Format.TextFrame2.TextRange.Characters(1, 7).Font.Bold = msoTrue
            Selection.Format.TextFrame2.TextRange.Characters(1, 7).Font.Size = axisFontSize
            Selection.Format.TextFrame2.TextRange.Characters(8, 2).Font.Bold = msoFalse
            Selection.Format.TextFrame2.TextRange.Characters(8, 2).Font.Size = dataLabelFontSize_secondary
        Else
            Selection.Format.TextFrame2.TextRange.Characters(1, 8).Font.Bold = msoTrue
            Selection.Format.TextFrame2.TextRange.Characters(1, 8).Font.Size = axisFontSize
            Selection.Format.TextFrame2.TextRange.Characters(10, 2).Font.Bold = msoFalse
            Selection.Format.TextFrame2.TextRange.Characters(10, 2).Font.Size = dataLabelFontSize_secondary
        End If

        cht.SeriesCollection(iseries).Points(2).DataLabel.Select
        Selection.Format.TextFrame2.TextRange.Font.Size = dataLabelFontSize_secondary
    Next iseries

    ' Squeeze plot area to make room for labels
    cht.PlotArea.Select
    Selection.Width = slopePlotWidth_web
    Selection.Left = slopePlotLeft
    Selection.Top = slopePlotTop

    If cht.HasLegend Then
        cht.Legend.Select
        Selection.Top = legend_top
        Selection.Left = legend_leftPad
        Selection.Font.Size = axisFontSize
    End If

    ' Remove border
    cht.ChartArea.Border.LineStyle = xlNone
End Sub


Sub SlopeChart()
    BuildSlopeChart
End Sub

Public Sub StyleSlopeChart_onAction(control As IRibbonControl)
    BuildSlopeChart
End Sub
