Attribute VB_Name = "modChartDotPlot"
Option Explicit

Private Sub BuildDotPlot()
    Dim inputValue As Variant
    Dim maxcell As Double
    Dim answer As Integer
    Dim ipts As Long
    Dim pointscount As Long
    Dim srs As Series
    Dim rng As String

    Dim seriescount As Long
    Dim imarker As Long
    Dim cht As Chart

    ' --- Prompt for group count ---
    inputValue = InputBox("How many groups (i.e., rows) do you want in your dot plot? (Must be at least 2 groups.)" & vbNewLine & vbCrLf & _
        "Styling slope charts will sometimes crash in Excel; if that occurs, simply delete the created chart and run it again.", "Input Box Text", "2")

    If inputValue = "" Then
        MsgBox "Dot Plot creation cancelled."
        Exit Sub
    End If
    If inputValue < 2 Then
        MsgBox "You have entered too few groups. Please run again with at least 2 groups.", vbExclamation
        Exit Sub
    End If

    maxcell = inputValue

    ' --- Build placeholder data and raw chart ---
    If inputValue > 1 Then

        Sheets.Add After:=ActiveSheet
        ActiveSheet.Name = WorksheetFunction.Text(Now(), "h_mm_ss")

        Range("A2").Select
        ActiveCell.FormulaR1C1 = "Group 1"
        Range("A2").Select
        Selection.AutoFill Destination:=Range("A2:A" & maxcell + 1), Type:=xlFillDefault

        Range("C1").Select
        ActiveCell.FormulaR1C1 = "Data Field A"
        Range("C2").Select
        ActiveCell.FormulaR1C1 = "20"
        Range("C3").Select
        ActiveCell.FormulaR1C1 = "=R[-1]C+10"
        Range("C3").Select
        If inputValue > 2 Then Selection.AutoFill Destination:=Range("C3:C" & maxcell + 1)

        Range("D1").Select
        ActiveCell.FormulaR1C1 = "Data Field B"
        Range("D2").Select
        ActiveCell.FormulaR1C1 = "30"
        Range("D3").Select
        ActiveCell.FormulaR1C1 = "=R[-1]C+10"
        Range("D3").Select
        If inputValue > 2 Then Selection.AutoFill Destination:=Range("D3:D" & maxcell + 1)

        Range("E1").Select
        ActiveCell.FormulaR1C1 = "Height"
        Range("E2").Select
        ActiveCell.FormulaR1C1 = maxcell * 2 - 1
        Range("E3").Select
        ActiveCell.FormulaR1C1 = "=R[-1]C-2"
        Range("E3").Select
        If inputValue > 2 Then Selection.AutoFill Destination:=Range("E3:E" & maxcell + 1)

        Range("F1").Select
        ActiveCell.FormulaR1C1 = "Error"
        Range("F2").Select
        ActiveCell.FormulaR1C1 = "=RC[-2]-RC[-3]"
        Range("F2").Select
        Selection.AutoFill Destination:=Range("F2:F" & maxcell + 1)

        Range("B1").Select
        ActiveCell.FormulaR1C1 = "Left Label"
        Range("B2").Select
        ActiveCell.FormulaR1C1 = "=RC[-1]&"" ""&RC[1]"
        Range("B2").Select
        Selection.AutoFill Destination:=Range("B2:B" & maxcell + 1)

        Range("C1:F1").Select
        With Selection
            .HorizontalAlignment = xlRight
            .VerticalAlignment = xlBottom
        End With
        Range("B1:B1").Select
        With Selection
            .HorizontalAlignment = xlLeft
            .VerticalAlignment = xlBottom
        End With

        Columns("B:B").EntireColumn.AutoFit
        Columns("C:C").EntireColumn.AutoFit
        Columns("D:D").EntireColumn.AutoFit

        Range("C1:D" & maxcell + 1).Select
        With Selection.Interior
            .Pattern = xlSolid
            .Color = giRGBbluecolor1
        End With

        Range("C2:C" & maxcell + 1, "D2:D" & maxcell + 1).Select
        ActiveSheet.Shapes.AddChart.Select
        ActiveChart.ChartType = xlXYScatter
        ActiveChart.SeriesCollection(2).XValues = Range("C2:C" & maxcell + 1)
        ActiveChart.SeriesCollection(2).Values = Range("E2:E" & maxcell + 1)
        ActiveChart.SeriesCollection(1).Delete

        ActiveChart.SeriesCollection.NewSeries
        ActiveChart.SeriesCollection(2).XValues = Range("D2:D" & maxcell + 1)
        ActiveChart.SeriesCollection(2).Values = Range("E2:E" & maxcell + 1)

        ActiveChart.FullSeriesCollection(1).Name = Range("C1")
        ActiveChart.FullSeriesCollection(2).Name = Range("D1")

        With ActiveChart
            .Parent.Top = dotPlotChartTop
            .Parent.Left = dotPlotChartLeft
        End With

        ActiveChart.PlotArea.Select
        ActiveChart.SeriesCollection(1).Select
        ActiveChart.SeriesCollection(1).ApplyDataLabels
        ActiveChart.SeriesCollection(1).DataLabels.Select
        Application.CutCopyMode = False

        rng = "B2:B" & maxcell + 1
        With ActiveChart.SeriesCollection(1)
            .ApplyDataLabels
            With .DataLabels
                ActiveChart.SeriesCollection(1).DataLabels.Format.TextFrame2.TextRange. _
                    InsertChartField msoChartFieldRange, _
                    "='" & ActiveSheet.Name & "'!" & rng, 0
                .ShowCategoryName = False
                .ShowRange = True
                .ShowSeriesName = False
                .ShowValue = False
                .Position = xlLabelPositionLeft
                .Separator = " "
                .Font.Size = dotPlotLabelFontSize
            End With
        End With

        For Each srs In ActiveChart.SeriesCollection
            With srs
                pointscount = .Points.Count
                For ipts = 1 To pointscount
                    ActiveChart.SeriesCollection(2).Points(ipts).ApplyDataLabels
                    With ActiveChart.SeriesCollection(2).Points(ipts).DataLabel
                        .Position = xlLabelPositionRight
                        .Font.Size = dotPlotLabelFontSize
                        .ShowSeriesName = False
                        .ShowValue = False
                        .ShowCategoryName = True
                        .Separator = " "
                    End With
                Next
            End With
        Next

        ActiveChart.Legend.Select
        ActiveChart.SetElement (msoElementLegendTop)

        ActiveChart.Axes(xlCategory).Select
        Selection.Delete

        ActiveChart.Axes(xlValue).MajorGridlines.Select
        Selection.Delete

        ActiveChart.Axes(xlValue).Select
        Selection.Delete

        ActiveChart.SeriesCollection(2).Select
        ActiveChart.SeriesCollection(2).HasErrorBars = True
        ActiveChart.SeriesCollection(2).ErrorBars.Select
        Selection.Delete
        ActiveChart.SeriesCollection(2).ErrorBar Direction:=xlX, Include:= _
            xlMinusValues, Type:=xlCustom, Amount:=Range("F2:F" & maxcell + 1), _
            MinusValues:=Range("F2:F" & maxcell + 1)
        ActiveChart.SeriesCollection(2).ErrorBars.EndStyle = xlNoCap

        ActiveChart.SeriesCollection(1).Select
        Selection.MarkerStyle = 8
        ActiveChart.SeriesCollection(2).Select
        Selection.MarkerStyle = 8

    End If ' End adding data

    ' --- Apply INSO formatting pipeline ---
    ActiveSheet.ChartObjects(1).Activate
    ActiveChart.Parent.Duplicate.Select

    Set cht = ActiveChart
    If cht Is Nothing Then Exit Sub

    ' Dot plot uses a partial pipeline (no gridlines, axis, or series color steps)
    OuterFormat cht
    FormatXAxisTitle cht
    InsertLogo cht
    InsertSource cht
    FormatTitle cht

    ' Dot plots don't use a Y-axis label
    On Error Resume Next
    cht.Shapes("YAxisLabelBox").Delete
    On Error GoTo 0

    DotPlotStyles cht

    ' Error bar color
    cht.SeriesCollection(2).ErrorBars.Select
    With Selection.Format.Line
        .Visible = msoTrue
        .ForeColor.rgb = giRGBbluecolor5
        .transparency = 0
    End With

    ' Marker style, size, and color
    seriescount = cht.SeriesCollection.Count
    For imarker = 1 To seriescount
        cht.SeriesCollection(imarker).Select
        With Selection
            .MarkerStyle = xlMarkerStyleCircle
            .MarkerSize = dotPlotMarkerSize
        End With
        With Selection.Format.Fill
            .Visible = msoTrue
            .ForeColor.rgb = giRGBbluecolor5
        End With
        With Selection
            .MarkerForegroundColorIndex = -4142
        End With
    Next imarker

    Range("A1").Select
End Sub


Sub DotPlot()
    BuildDotPlot
End Sub

Public Sub StyleDotPlot_onAction(control As IRibbonControl)
    BuildDotPlot
End Sub
