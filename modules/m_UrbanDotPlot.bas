Attribute VB_Name = "m_UrbanDotPlot"
Option Explicit



Sub UrbanDotPlot()
'
    Dim inputValue As Variant
    Dim maxcell As Double
    Dim answer As Integer
    Dim ipts As Long
    Dim pointscount As Long
    Dim srs As Series
    Dim rng As String
    
    'Variables to create the styled graph
    Dim seriescount As Long
    Dim iseries As Long
    Dim imarker As Long
    Dim cht As Chart
    Dim txtB As TextBox
    Dim bInsertLogo As Boolean
    Dim bInsertSource As Boolean
    Dim bFormatTitle As Boolean
    Dim bFormatGridlines As Boolean
    Dim bFormatXAxis As Boolean
    Dim bFormatXAxisTitle As Boolean
    Dim bFormatLineColors As Boolean
    Dim bOuterFormat As Boolean
    Dim bSetWebVersion As Boolean
    Dim bDotPlotStyles As Boolean
    
    
    'Create placeholder data for dot plot
        'Message box for input value
    inputValue = InputBox("How many groups (i.e., rows) do you want in your dot plot? (Must be at least 2 groups.)" & vbNewLine & vbCrLf & _
    "Styling slope charts will sometimes crash in Excel; if that occurs, simply delete the created chart and run it again.", "Input Box Text", "2")

    If inputValue = "" Then
        MsgBox "Dot Plot creation cancelled."
        Exit Sub
    End If
    If inputValue < 2 Then
        answer = MsgBox("You have entered too few groups. Please run again with at least 2 groups.", vbExclamation)
        Exit Sub
    End If
''    If inputValue > 10 Then
''        answer = MsgBox("You have entered too many groups. Please run again with between 2 and 10 groups.", vbExclamation)
''        Exit Sub
''    End If

    maxcell = inputValue

    'Add the data
    If inputValue > 1 Then
    
        'Create new sheet
        'On Error GoTo MyError
        Sheets.Add After:=ActiveSheet
        ActiveSheet.name = WorksheetFunction.Text(Now(), "h_mm_ss")

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
        If inputValue > 2 Then
            Selection.AutoFill Destination:=Range("C3:C" & maxcell + 1)
        End If

        Range("D1").Select
        ActiveCell.FormulaR1C1 = "Data Field B"
        Range("D2").Select
        ActiveCell.FormulaR1C1 = "30"
        Range("D3").Select
        ActiveCell.FormulaR1C1 = "=R[-1]C+10"
        Range("D3").Select
        If inputValue > 2 Then
            Selection.AutoFill Destination:=Range("D3:D" & maxcell + 1)
        End If

        Range("E1").Select
        ActiveCell.FormulaR1C1 = "Height"
        Range("E2").Select
        ActiveCell.FormulaR1C1 = maxcell * 2 - 1
        Range("E3").Select
        ActiveCell.FormulaR1C1 = "=R[-1]C-2"
        Range("E3").Select
        If inputValue > 2 Then
            Selection.AutoFill Destination:=Range("E3:E" & maxcell + 1)
        End If

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

        'Align data column header names
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
        
        'Autofit data columns
        Columns("B:B").EntireColumn.EntireColumn.AutoFit
        Columns("C:C").EntireColumn.EntireColumn.AutoFit
        Columns("D:D").EntireColumn.EntireColumn.AutoFit

        'Fill data to be filled in
         Range("C1:D" & maxcell + 1).Select
         With Selection.Interior
            .Pattern = xlSolid
            .color = rgb(207, 232, 243)
         End With

        Range("C2:C" & maxcell + 1, "D2:D" & maxcell + 1).Select
        ActiveSheet.Shapes.AddChart.Select
        ActiveChart.chartType = xlXYScatter
        ActiveChart.SeriesCollection(2).XValues = Range("C2:C" & maxcell + 1)
        ActiveChart.SeriesCollection(2).Values = Range("E2:E" & maxcell + 1)
        ActiveChart.SeriesCollection(1).Delete

        ActiveChart.SeriesCollection.NewSeries
        ActiveChart.SeriesCollection(2).XValues = Range("D2:D" & maxcell + 1)
        ActiveChart.SeriesCollection(2).Values = Range("E2:E" & maxcell + 1)

        ActiveChart.FullSeriesCollection(1).name = Range("C1")
        ActiveChart.FullSeriesCollection(2).name = Range("D1")

        'Move chart
        With ActiveChart
            .Parent.Top = 10
            .Parent.Left = 350
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
            "='" & ActiveSheet.name & "'!" & rng, 0
            .ShowCategoryName = False
            .ShowRange = True
            .ShowSeriesName = False
            .ShowValue = False
            .Position = xlLabelPositionLeft
            .Separator = " "
            .Font.Size = 8
        End With
    End With
        
        'Add data labels to the right-most point
        'Add and format data labels
        'Left (first) point
        For Each srs In ActiveChart.SeriesCollection
            With srs
                pointscount = .Points.Count
                For ipts = 1 To pointscount
                    ActiveChart.SeriesCollection(2).Points(ipts).ApplyDataLabels
                    With ActiveChart.SeriesCollection(2).Points(ipts).DataLabel
                        .Position = xlLabelPositionRight
                        .Font.Size = 8
                        .ShowSeriesName = False
                        .ShowValue = False
                        .ShowCategoryName = True
                        .Separator = " "
                    End With
                Next
            End With
        Next

        'Move legend to top
        ActiveChart.Legend.Select
        ActiveChart.SetElement (msoElementLegendTop)

''        'Special case of 2 rows, need to change the maximum of the secondary axis
''        If inputValue = 2 Then
''            ActiveChart.Axes(xlValue, xlSecondary).Select
''            ActiveChart.Axes(xlValue, xlSecondary).MaximumScale = 4
''        End If
        
        'Delete both vertical axes
        ActiveChart.Axes(xlCategory).Select
        Selection.Delete

        'Delete vertical gridlines
        ActiveChart.Axes(xlValue).MajorGridlines.Select
        Selection.Delete

        'Delete x-axis labels
        ActiveChart.Axes(xlValue).Select
        Selection.Delete

        'Add error bars
        ActiveChart.SeriesCollection(2).Select
        ActiveChart.SeriesCollection(2).HasErrorBars = True
        ActiveChart.SeriesCollection(2).ErrorBars.Select
        Selection.Delete
        ActiveChart.SeriesCollection(2).ErrorBar Direction:=xlX, Include:= _
                xlMinusValues, Type:=xlCustom, Amount:=Range("F2:F" & maxcell + 1), _
                MinusValues:=Range("F2:F" & maxcell + 1)
        ActiveChart.SeriesCollection(2).ErrorBars.EndStyle = xlNoCap

        'Color the dots
        ActiveChart.SeriesCollection(1).Select
        Selection.MarkerStyle = 8
        ActiveChart.SeriesCollection(2).Select
        Selection.MarkerStyle = 8

    End If 'End Adding data

'''''''   ActiveChart.ChartArea.Select

    
''''''    'Check ActiveChart
''''''    If ActiveChart Is Nothing Then
''''''        MsgBox "First, create a chart, or select an existing chart, and try again.", vbExclamation, "No Active Chart"
''''''    Else
    
        ''''bSetWebVersion = SetWebVersion(gWebCancel)
        SetWebVersion_NEW
        If gWebCancel Then
            Exit Sub
        End If

        'Select the (only) chart in the worksheet
        ActiveSheet.ChartObjects(1).Activate
        
        'Duplicate the selected chart
        ActiveChart.Parent.Duplicate.Select

        Set cht = ActiveChart

        bOuterFormat = OuterFormat(cht)
        bFormatXAxisTitle = FormatXAxisTitle(cht)
        bInsertLogo = InsertLogo(cht)
        bInsertSource = InsertSource(cht)
        bFormatTitle = FormatTitle(cht)
            'Dot plots don't need y axis labels
            cht.Shapes.Range(Array("YAxisLabelBox")).Select
            Selection.Delete
            
        'bFormatGridlines = FormatGridlines(cht)
        'bFormatXAxis = FormatXAxis(cht)
        'bFormatLineColors = FormatSeriesColors(cht, "LINE")
        bDotPlotStyles = DotPlotStyles(cht)
        
        'Change color of error bars
        ActiveChart.SeriesCollection(2).ErrorBars.Select
        With Selection.Format.Line
            .Visible = msoTrue
            .ForeColor.ObjectThemeColor = msoThemeColorAccent1
            .ForeColor.TintAndShade = 0
            .ForeColor.Brightness = 0
        End With
        With Selection.Format.Line
            .Visible = msoTrue
            .ForeColor.rgb = giRGBbluecolor5
            .transparency = 0
        End With

        'Change marker style and size
        With ActiveChart
            seriescount = .SeriesCollection.Count
        End With

        For imarker = 1 To seriescount
            ActiveChart.SeriesCollection(imarker).Select
            With Selection
                .MarkerStyle = xlMarkerStyleCircle
                .MarkerSize = 6
            End With
            With Selection.Format.Fill
                .Visible = msoTrue
                .ForeColor.rgb = giRGBbluecolor5
            End With
            With Selection
                '.Marker.Line.Visible = msoFalse
                .MarkerForegroundColorIndex = -4142
            End With
        Next

''''''    End If    'Active Chart message box

    'Put the cursor into cell A1
    Range("A1").Select
    
End Sub


Public Sub StyleDotPlot_onAction(control As IRibbonControl)
'
' Code for onAction callback. Ribbon control button
'
    'Variables to create the placeholder data
    Dim inputValue As Variant
    Dim maxcell As Double
    Dim answer As Integer
    Dim ipts As Long
    Dim pointscount As Long
    Dim srs As Series
    Dim rng As String
    
    'Variables to create the styled graph
    Dim seriescount As Long
    Dim iseries As Long
    Dim imarker As Long
    Dim cht As Chart
    Dim txtB As TextBox
    Dim bInsertLogo As Boolean
    Dim bInsertSource As Boolean
    Dim bFormatTitle As Boolean
    Dim bFormatGridlines As Boolean
    Dim bFormatXAxis As Boolean
    Dim bFormatXAxisTitle As Boolean
    Dim bFormatLineColors As Boolean
    Dim bOuterFormat As Boolean
    Dim bSetWebVersion As Boolean
    Dim bDotPlotStyles As Boolean
    
    
    'Create placeholder data for dot plot
        'Message box for input value
    inputValue = InputBox("How many groups (i.e., rows) do you want in your dot plot? (Must be at least 2 groups.)" & vbNewLine & vbCrLf & _
    "Styling slope charts will sometimes crash in Excel; if that occurs, simply delete the created chart and run it again.", "Input Box Text", "2")

    If inputValue = "" Then
        MsgBox "Dot Plot creation cancelled."
        Exit Sub
    End If
    If inputValue < 2 Then
        answer = MsgBox("You have entered too few groups. Please run again with at least 2 groups.", vbExclamation)
        Exit Sub
    End If
''    If inputValue > 10 Then
''        answer = MsgBox("You have entered too many groups. Please run again with between 2 and 10 groups.", vbExclamation)
''        Exit Sub
''    End If

    maxcell = inputValue

    'Add the data
    If inputValue > 1 Then
    
        'Create new sheet
        'On Error GoTo MyError
        Sheets.Add After:=ActiveSheet
        ActiveSheet.name = WorksheetFunction.Text(Now(), "h_mm_ss")

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
        If inputValue > 2 Then
            Selection.AutoFill Destination:=Range("C3:C" & maxcell + 1)
        End If

        Range("D1").Select
        ActiveCell.FormulaR1C1 = "Data Field B"
        Range("D2").Select
        ActiveCell.FormulaR1C1 = "30"
        Range("D3").Select
        ActiveCell.FormulaR1C1 = "=R[-1]C+10"
        Range("D3").Select
        If inputValue > 2 Then
            Selection.AutoFill Destination:=Range("D3:D" & maxcell + 1)
        End If

        Range("E1").Select
        ActiveCell.FormulaR1C1 = "Height"
        Range("E2").Select
        ActiveCell.FormulaR1C1 = maxcell * 2 - 1
        Range("E3").Select
        ActiveCell.FormulaR1C1 = "=R[-1]C-2"
        Range("E3").Select
        If inputValue > 2 Then
            Selection.AutoFill Destination:=Range("E3:E" & maxcell + 1)
        End If

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

        'Align data column header names
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
        
        'Autofit data columns
        Columns("B:B").EntireColumn.EntireColumn.AutoFit
        Columns("C:C").EntireColumn.EntireColumn.AutoFit
        Columns("D:D").EntireColumn.EntireColumn.AutoFit

        'Fill data to be filled in
         Range("C1:D" & maxcell + 1).Select
         With Selection.Interior
            .Pattern = xlSolid
            .color = rgb(207, 232, 243)
         End With

        Range("C2:C" & maxcell + 1, "D2:D" & maxcell + 1).Select
        ActiveSheet.Shapes.AddChart.Select
        ActiveChart.chartType = xlXYScatter
        ActiveChart.SeriesCollection(2).XValues = Range("C2:C" & maxcell + 1)
        ActiveChart.SeriesCollection(2).Values = Range("E2:E" & maxcell + 1)
        ActiveChart.SeriesCollection(1).Delete

        ActiveChart.SeriesCollection.NewSeries
        ActiveChart.SeriesCollection(2).XValues = Range("D2:D" & maxcell + 1)
        ActiveChart.SeriesCollection(2).Values = Range("E2:E" & maxcell + 1)

        ActiveChart.FullSeriesCollection(1).name = Range("C1")
        ActiveChart.FullSeriesCollection(2).name = Range("D1")

        'Move chart
        With ActiveChart
            .Parent.Top = 10
            .Parent.Left = 350
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
            "='" & ActiveSheet.name & "'!" & rng, 0
            .ShowCategoryName = False
            .ShowRange = True
            .ShowSeriesName = False
            .ShowValue = False
            .Position = xlLabelPositionLeft
            .Separator = " "
            .Font.Size = 8
        End With
    End With
        
        'Add data labels to the right-most point
        'Add and format data labels
        'Left (first) point
        For Each srs In ActiveChart.SeriesCollection
            With srs
                pointscount = .Points.Count
                For ipts = 1 To pointscount
                    ActiveChart.SeriesCollection(2).Points(ipts).ApplyDataLabels
                    With ActiveChart.SeriesCollection(2).Points(ipts).DataLabel
                        .Position = xlLabelPositionRight
                        .Font.Size = 8
                        .ShowSeriesName = False
                        .ShowValue = False
                        .ShowCategoryName = True
                        .Separator = " "
                    End With
                Next
            End With
        Next

        'Move legend to top
        ActiveChart.Legend.Select
        ActiveChart.SetElement (msoElementLegendTop)

''        'Special case of 2 rows, need to change the maximum of the secondary axis
''        If inputValue = 2 Then
''            ActiveChart.Axes(xlValue, xlSecondary).Select
''            ActiveChart.Axes(xlValue, xlSecondary).MaximumScale = 4
''        End If
        
        'Delete both vertical axes
        ActiveChart.Axes(xlCategory).Select
        Selection.Delete

        'Delete vertical gridlines
        ActiveChart.Axes(xlValue).MajorGridlines.Select
        Selection.Delete

        'Delete x-axis labels
        ActiveChart.Axes(xlValue).Select
        Selection.Delete

        'Add error bars
        ActiveChart.SeriesCollection(2).Select
        ActiveChart.SeriesCollection(2).HasErrorBars = True
        ActiveChart.SeriesCollection(2).ErrorBars.Select
        Selection.Delete
        ActiveChart.SeriesCollection(2).ErrorBar Direction:=xlX, Include:= _
                xlMinusValues, Type:=xlCustom, Amount:=Range("F2:F" & maxcell + 1), _
                MinusValues:=Range("F2:F" & maxcell + 1)
        ActiveChart.SeriesCollection(2).ErrorBars.EndStyle = xlNoCap

        'Color the dots
        ActiveChart.SeriesCollection(1).Select
        Selection.MarkerStyle = 8
        ActiveChart.SeriesCollection(2).Select
        Selection.MarkerStyle = 8

    End If 'End Adding data

'''''''   ActiveChart.ChartArea.Select

    
''''''    'Check ActiveChart
''''''    If ActiveChart Is Nothing Then
''''''        MsgBox "First, create a chart, or select an existing chart, and try again.", vbExclamation, "No Active Chart"
''''''    Else
    
        ''''bSetWebVersion = SetWebVersion(gWebCancel)
        SetWebVersion_NEW
        If gWebCancel Then
            Exit Sub
        End If

        'Select the (only) chart in the worksheet
        ActiveSheet.ChartObjects(1).Activate
        
        'Duplicate the selected chart
        ActiveChart.Parent.Duplicate.Select

        Set cht = ActiveChart

        bOuterFormat = OuterFormat(cht)
        bFormatXAxisTitle = FormatXAxisTitle(cht)
        bInsertLogo = InsertLogo(cht)
        bInsertSource = InsertSource(cht)
        bFormatTitle = FormatTitle(cht)
            'Dot plots don't need y axis labels
            cht.Shapes.Range(Array("YAxisLabelBox")).Select
            Selection.Delete
            
        'bFormatGridlines = FormatGridlines(cht)
        'bFormatXAxis = FormatXAxis(cht)
        'bFormatLineColors = FormatSeriesColors(cht, "LINE")
        bDotPlotStyles = DotPlotStyles(cht)
        
        'Change color of error bars
        ActiveChart.SeriesCollection(2).ErrorBars.Select
        With Selection.Format.Line
            .Visible = msoTrue
            .ForeColor.ObjectThemeColor = msoThemeColorAccent1
            .ForeColor.TintAndShade = 0
            .ForeColor.Brightness = 0
        End With
        With Selection.Format.Line
            .Visible = msoTrue
            .ForeColor.rgb = giRGBbluecolor5
            .transparency = 0
        End With

        'Change marker style and size
        With ActiveChart
            seriescount = .SeriesCollection.Count
        End With

        For imarker = 1 To seriescount
            ActiveChart.SeriesCollection(imarker).Select
            With Selection
                .MarkerStyle = xlMarkerStyleCircle
                .MarkerSize = 6
            End With
            With Selection.Format.Fill
                .Visible = msoTrue
                .ForeColor.rgb = giRGBbluecolor5
            End With
            With Selection
                '.Marker.Line.Visible = msoFalse
                .MarkerForegroundColorIndex = -4142
            End With
        Next

''''''    End If    'Active Chart message box

    'Put the cursor into cell A1
    Range("A1").Select

End Sub

'JAS 2017
'JAS 2023
