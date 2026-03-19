Attribute VB_Name = "m_UrbanSlopeChart"

Option Explicit



Sub UrbanSlopeChart()
'
    'Variables for data creation
    Dim inputValue As Variant
    Dim maxcell As Double
    Dim bChartType As String
    Dim answer As Integer

    'Variables for slope chart styling
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
    Dim bSlopeChartStyles As Boolean

    inputValue = InputBox("How many groups (i.e., rows) do you want in your slope chart? (Must be at least 2 groups.)" & vbNewLine & vbCrLf & _
    "Styling slope charts will sometimes crash in Excel; if that occurs, simply delete the created chart and run it again." & vbNewLine & vbCrLf & _
    "Note that for slope charts with more than 6 series, those series will be assigned a gray color.", "Input Box Text", "2")

    If inputValue = "" Then
        MsgBox "Slope Chart creation cancelled."
        Exit Sub
    End If
    If inputValue < 2 Then
        answer = MsgBox("You have entered too few groups. Please run again with at least 2 groups.", vbExclamation)
        Exit Sub
    End If
''    If inputValue > 6 Then
''        answer = MsgBox("You have entered too many groups. Please run again with between 2 and 6 groups.", vbExclamation)
''        Exit Sub
''    End If

    maxcell = inputValue

''    If inputValue > 1 And inputValue <= 6 Then
    If inputValue > 1 Then
    
        'Create new sheet
        'On Error GoTo MyError
        Sheets.Add After:=ActiveSheet
        ActiveSheet.name = WorksheetFunction.Text(Now(), "h_mm_ss")

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


        'Create the slope chart
        bChartType = xlLine
        Range("A1:C" & maxcell + 1).Select
        '   ActiveSheet.Shapes.AddChart2(227, bChartType, 50, 125).Select
        ActiveSheet.Shapes.AddChart.Select
        ActiveChart.chartType = bChartType
        ActiveChart.SetSourceData Source:=Range("A1:C" & maxcell + 1)

        'Switch rows and columns
        ActiveChart.PlotBy = xlRows
        'Extend x-axis to go to tick marks
        ActiveChart.Axes(xlCategory).Select
        ActiveChart.Axes(xlCategory).AxisBetweenCategories = False
        'Delete gridlines
        ActiveChart.Axes(xlValue).MajorGridlines.Select
        Selection.Delete
        'Delete y-axis
        ActiveChart.Axes(xlValue).Select
        Selection.Delete

        'Move chart
        With ActiveChart
            .Parent.Top = 10
            .Parent.Left = 350
        End With

    End If
    
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
        bFormatGridlines = FormatGridlines(cht)
        'bFormatXAxis = FormatXAxis(cht)
        bFormatLineColors = FormatSeriesColors(cht, "LINE")

        'Position xaxis on tick marks
        ActiveChart.Axes(xlCategory).AxisBetweenCategories = False

        'Style xaxis and yaxis tick marks
        ActiveChart.Axes(xlCategory).MajorTickMark = xlTickMarkOutside
        ActiveChart.Axes(xlCategory).MinorTickMark = xlTickMarkNone
        ActiveChart.Axes(xlValue).MajorTickMark = xlTickMarkOutside
        ActiveChart.Axes(xlValue).MinorTickMark = xlTickMarkNone

        'Remove Gridlines
        If ActiveChart.Axes(xlValue).HasMajorGridlines = True Then
            ActiveChart.Axes(xlValue).MajorGridlines.Select
            Selection.Delete
        End If

        'Change x-axis line color
        cht.Axes(xlCategory).Select
        With Selection.Format.Line
            .Visible = msoTrue
            '.ForeColor.ObjectThemeColor = msoThemeColorText1
            .ForeColor.rgb = colorBlack
            .ForeColor.TintAndShade = 0
            .ForeColor.Brightness = 0
            .weight = 1
        End With

        With ActiveChart
            seriescount = .SeriesCollection.Count
        End With

        'Add and format data labels
        'Left (first) point
        For iseries = 1 To seriescount
            ActiveChart.SeriesCollection(iseries).Points(1).ApplyDataLabels
            With ActiveChart.SeriesCollection(iseries).Points(1).DataLabel
                .Position = xlLabelPositionLeft
                .Font.Size = 8
                .ShowSeriesName = True
                .Separator = " "
            End With
        Next

        'Right (last) point
        For iseries = 1 To seriescount
            ActiveChart.SeriesCollection(iseries).Points(2).ApplyDataLabels
            With ActiveChart.SeriesCollection(iseries).Points(2).DataLabel
                .Position = xlLabelPositionRight
                .Font.Size = 8
                .ShowSeriesName = False
                .Separator = " "
            End With
        Next

        'Change marker style and size
        For imarker = 1 To seriescount
            ActiveChart.SeriesCollection(imarker).Select
            With Selection
                .MarkerStyle = xlMarkerStyleCircle
                .MarkerSize = 7
            End With
            With Selection.Format.Fill
                .Visible = msoTrue
                .ForeColor.rgb = giRGBwhitecolor
            End With
        Next

        'Final formatting issues--format labels, move plot area over, and legend down slightly
        bSlopeChartStyles = SlopeChartStyles(cht)

''''''    End If    'Active Chart message box

    'Put the cursor into cell A1
    Range("A1").Select

End Sub

Public Sub StyleSlopeChart_onAction(control As IRibbonControl)
'
' Code for onAction callback. Ribbon control button
'

    'Variables for data creation
    Dim inputValue As Variant
    Dim maxcell As Double
    Dim bChartType As String
    Dim answer As Integer

    'Variables for slope chart styling
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
    Dim bSlopeChartStyles As Boolean

    inputValue = InputBox("How many groups (i.e., rows) do you want in your slope chart? (Must be at least 2 groups.)" & vbNewLine & vbCrLf & _
    "Styling slope charts will sometimes crash in Excel; if that occurs, simply delete the created chart and run it again." & vbNewLine & vbCrLf & _
    "Note that for slope charts with more than 6 series, those series will be assigned a gray color.", "Input Box Text", "2")

    If inputValue = "" Then
        MsgBox "Slope Chart creation cancelled."
        Exit Sub
    End If
    If inputValue < 2 Then
        answer = MsgBox("You have entered too few groups. Please run again with at least 2 groups.", vbExclamation)
        Exit Sub
    End If
''    If inputValue > 6 Then
''        answer = MsgBox("You have entered too many groups. Please run again with between 2 and 6 groups.", vbExclamation)
''        Exit Sub
''    End If

    maxcell = inputValue

''    If inputValue > 1 And inputValue <= 6 Then
    If inputValue > 1 Then
    
        'Create new sheet
        'On Error GoTo MyError
        Sheets.Add After:=ActiveSheet
        ActiveSheet.name = WorksheetFunction.Text(Now(), "h_mm_ss")

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


        'Create the slope chart
        bChartType = xlLine
        Range("A1:C" & maxcell + 1).Select
        '   ActiveSheet.Shapes.AddChart2(227, bChartType, 50, 125).Select
        ActiveSheet.Shapes.AddChart.Select
        ActiveChart.chartType = bChartType
        ActiveChart.SetSourceData Source:=Range("A1:C" & maxcell + 1)

        'Switch rows and columns
        ActiveChart.PlotBy = xlRows
        'Extend x-axis to go to tick marks
        ActiveChart.Axes(xlCategory).Select
        ActiveChart.Axes(xlCategory).AxisBetweenCategories = False
        'Delete gridlines
        ActiveChart.Axes(xlValue).MajorGridlines.Select
        Selection.Delete
        'Delete y-axis
        ActiveChart.Axes(xlValue).Select
        Selection.Delete

        'Move chart
        With ActiveChart
            .Parent.Top = 10
            .Parent.Left = 350
        End With

    End If
    
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
        bFormatGridlines = FormatGridlines(cht)
        'bFormatXAxis = FormatXAxis(cht)
        bFormatLineColors = FormatSeriesColors(cht, "LINE")

        'Position xaxis on tick marks
        ActiveChart.Axes(xlCategory).AxisBetweenCategories = False

        'Style xaxis and yaxis tick marks
        ActiveChart.Axes(xlCategory).MajorTickMark = xlTickMarkOutside
        ActiveChart.Axes(xlCategory).MinorTickMark = xlTickMarkNone
        ActiveChart.Axes(xlValue).MajorTickMark = xlTickMarkOutside
        ActiveChart.Axes(xlValue).MinorTickMark = xlTickMarkNone

        'Remove Gridlines
        If ActiveChart.Axes(xlValue).HasMajorGridlines = True Then
            ActiveChart.Axes(xlValue).MajorGridlines.Select
            Selection.Delete
        End If

        'Change x-axis line color
        cht.Axes(xlCategory).Select
        With Selection.Format.Line
            .Visible = msoTrue
            '.ForeColor.ObjectThemeColor = msoThemeColorText1
            .ForeColor.rgb = colorBlack
            .ForeColor.TintAndShade = 0
            .ForeColor.Brightness = 0
            .weight = 1
        End With

        With ActiveChart
            seriescount = .SeriesCollection.Count
        End With

        'Add and format data labels
        'Left (first) point
        For iseries = 1 To seriescount
            ActiveChart.SeriesCollection(iseries).Points(1).ApplyDataLabels
            With ActiveChart.SeriesCollection(iseries).Points(1).DataLabel
                .Position = xlLabelPositionLeft
                .Font.Size = 8
                .ShowSeriesName = True
                .Separator = " "
            End With
        Next

        'Right (last) point
        For iseries = 1 To seriescount
            ActiveChart.SeriesCollection(iseries).Points(2).ApplyDataLabels
            With ActiveChart.SeriesCollection(iseries).Points(2).DataLabel
                .Position = xlLabelPositionRight
                .Font.Size = 8
                .ShowSeriesName = False
                .Separator = " "
            End With
        Next

        'Change marker style and size
        For imarker = 1 To seriescount
            ActiveChart.SeriesCollection(imarker).Select
            With Selection
                .MarkerStyle = xlMarkerStyleCircle
                .MarkerSize = 7
            End With
            With Selection.Format.Fill
                .Visible = msoTrue
                .ForeColor.rgb = giRGBwhitecolor
            End With
        Next

        'Final formatting issues--format labels, move plot area over, and legend down slightly
        bSlopeChartStyles = SlopeChartStyles(cht)

''''''    End If    'Active Chart message box

    'Put the cursor into cell A1
    Range("A1").Select
    
End Sub

'JAS 2017
'JAS 2023
