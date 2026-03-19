Attribute VB_Name = "m_UrbanScatterplot"
Option Explicit



Sub UrbanScatterplot()
'
    Dim seriescount As Long
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
    Dim bScatterplotStyles As Boolean

'    'Check ActiveChart
'    If ActiveChart Is Nothing Then
'        MsgBox "First, create a chart, or select an existing chart, and try again.", vbExclamation, "No Active Chart"
'    Else

        '''bSetWebVersion = SetWebVersion(gWebCancel)
        SetWebVersion_NEW
        If gWebCancel Then
            Exit Sub
        End If

        ActiveSheet.Shapes.AddChart2(-1, xlXYScatter).Select
        'Duplicate the selected chart
        ActiveChart.Parent.Duplicate.Select

        Set cht = ActiveChart

        bOuterFormat = OuterFormat(cht)
        bFormatXAxisTitle = FormatXAxisTitle(cht)
        bInsertLogo = InsertLogo(cht)
        bInsertSource = InsertSource(cht)
        bFormatTitle = FormatTitle(cht)
        'bFormatGridlines = FormatGridlines(cht)
        bFormatXAxis = FormatXAxis(cht)
        'bFormatLineColors = FormatSeriesColors(cht, "LINE")
        bScatterplotStyles = ScatterplotStyles(cht) 'formats gridlines

        'Style xaxis tick marks
        cht.Axes(xlCategory).MajorTickMark = xlTickMarkOutside
        cht.Axes(xlCategory).MinorTickMark = xlTickMarkNone

        'Change colors of points
        With cht
            seriescount = .SeriesCollection.Count
        End With

        'Change marker style and size
        For imarker = 1 To seriescount
            cht.SeriesCollection(imarker).Select
            With Selection
                .MarkerStyle = xlMarkerStyleCircle
                .MarkerSize = 7
            End With
        Next

        If seriescount > 0 And seriescount <= 6 Then
            With cht.SeriesCollection(1)
                .Border.ColorIndex = xlNone
                'ForegroundColorIndex lets you set line to no color
                .MarkerForegroundColorIndex = xlColorIndexNone
                'BackgroundColor is marker fill
                .MarkerBackgroundColor = giRGBbluecolor5
            End With
            If seriescount >= 2 Then
                With cht.SeriesCollection(2)
                    .Border.ColorIndex = xlNone
                    'ForegroundColorIndex lets you set line to no color
                    .MarkerForegroundColorIndex = xlColorIndexNone
                    'BackgroundColor is marker fill
                    .MarkerBackgroundColor = giRGByellowcolor
                End With
            End If
            If seriescount >= 3 Then
                With cht.SeriesCollection(3)
                    .Border.ColorIndex = xlNone
                    'ForegroundColorIndex lets you set line to no color
                    .MarkerForegroundColorIndex = xlColorIndexNone
                    'BackgroundColor is marker fill
                    .MarkerBackgroundColor = colorBlack
                End With
            End If
            If seriescount >= 4 Then
                With cht.SeriesCollection(4)
                    .Border.ColorIndex = xlNone
                    'ForegroundColorIndex lets you set line to no color
                    .MarkerForegroundColorIndex = xlColorIndexNone
                    'BackgroundColor is marker fill
                    .MarkerBackgroundColor = colorSilver
                End With
            End If
            If seriescount >= 5 Then
                With cht.SeriesCollection(5)
                    .Border.ColorIndex = xlNone
                    'ForegroundColorIndex lets you set line to no color
                    .MarkerForegroundColorIndex = xlColorIndexNone
                    'BackgroundColor is marker fill
                    .MarkerBackgroundColor = giRGBpinkcolor
                End With
            End If
            If seriescount >= 6 Then
                With cht.SeriesCollection(6)
                    .Border.ColorIndex = xlNone
                    'ForegroundColorIndex lets you set line to no color
                    .MarkerForegroundColorIndex = xlColorIndexNone
                    'BackgroundColor is marker fill
                    .MarkerBackgroundColor = giRGBgreencolor
                End With
            End If

        ElseIf seriescount > 6 Then
            If cht.HasTitle = True Then
                cht.ChartTitle.Text = "You have too many data series for this chart type. Please contact the Communications Department for further guidance."
            ElseIf cht.HasTitle = False Then
                'Add text box with text, size, and use correct font
                Set txtB = cht.TextBoxes.Add(0, 0, 500, 40)
                'Set txtB = cht.TextBoxes.Add(400, 100, 125, 20)
                '(horizontal placement, vertical placement, box width, box height)
                With txtB
                    .name = "TitleBox"
                    .Text = "You have too many data series for this chart type. Please contact the Communications Department for further guidance."
                    .Font.Size = 10
                    .Font.name = "Lato"
                End With
                'Left-align text in text box
                cht.Shapes.Range(Array("TitleBox")).TextEffect.Alignment = msoTextEffectAlignmentLeft
            End If

        End If

'    End If

End Sub

Public Sub Scatter_onAction(control As IRibbonControl)
'
' Code for onAction callback. Ribbon control button
'
'
    Dim seriescount As Long
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
    Dim bScatterplotStyles As Boolean
    
'    'Check ActiveChart
'    If ActiveChart Is Nothing Then
'        MsgBox "First, create a chart, or select an existing chart, and try again.", vbExclamation, "No Active Chart"
'    Else

        '''bSetWebVersion = SetWebVersion(gWebCancel)
        SetWebVersion_NEW
        If gWebCancel Then
            Exit Sub
        End If

        ActiveSheet.Shapes.AddChart2(-1, xlXYScatter).Select
        'Duplicate the selected chart
        ActiveChart.Parent.Duplicate.Select

        Set cht = ActiveChart

        bOuterFormat = OuterFormat(cht)
        bFormatXAxisTitle = FormatXAxisTitle(cht)
        bInsertLogo = InsertLogo(cht)
        bInsertSource = InsertSource(cht)
        bFormatTitle = FormatTitle(cht)
        'bFormatGridlines = FormatGridlines(cht)
        bFormatXAxis = FormatXAxis(cht)
        'bFormatLineColors = FormatSeriesColors(cht, "LINE")
        bScatterplotStyles = ScatterplotStyles(cht) 'formats gridlines

        'Style xaxis tick marks
        cht.Axes(xlCategory).MajorTickMark = xlTickMarkOutside
        cht.Axes(xlCategory).MinorTickMark = xlTickMarkNone

        'Change colors of points
        With cht
            seriescount = .SeriesCollection.Count
        End With

        'Change marker style and size
        For imarker = 1 To seriescount
            cht.SeriesCollection(imarker).Select
            With Selection
                .MarkerStyle = xlMarkerStyleCircle
                .MarkerSize = 7
            End With
        Next

        If seriescount > 0 And seriescount <= 6 Then
            With cht.SeriesCollection(1)
                .Border.ColorIndex = xlNone
                'ForegroundColorIndex lets you set line to no color
                .MarkerForegroundColorIndex = xlColorIndexNone
                'BackgroundColor is marker fill
                .MarkerBackgroundColor = giRGBbluecolor5
            End With
            If seriescount >= 2 Then
                With cht.SeriesCollection(2)
                    .Border.ColorIndex = xlNone
                    'ForegroundColorIndex lets you set line to no color
                    .MarkerForegroundColorIndex = xlColorIndexNone
                    'BackgroundColor is marker fill
                    .MarkerBackgroundColor = giRGByellowcolor
                End With
            End If
            If seriescount >= 3 Then
                With cht.SeriesCollection(3)
                    .Border.ColorIndex = xlNone
                    'ForegroundColorIndex lets you set line to no color
                    .MarkerForegroundColorIndex = xlColorIndexNone
                    'BackgroundColor is marker fill
                    .MarkerBackgroundColor = colorBlack
                End With
            End If
            If seriescount >= 4 Then
                With cht.SeriesCollection(4)
                    .Border.ColorIndex = xlNone
                    'ForegroundColorIndex lets you set line to no color
                    .MarkerForegroundColorIndex = xlColorIndexNone
                    'BackgroundColor is marker fill
                    .MarkerBackgroundColor = colorSilver
                End With
            End If
            If seriescount >= 5 Then
                With cht.SeriesCollection(5)
                    .Border.ColorIndex = xlNone
                    'ForegroundColorIndex lets you set line to no color
                    .MarkerForegroundColorIndex = xlColorIndexNone
                    'BackgroundColor is marker fill
                    .MarkerBackgroundColor = giRGBpinkcolor
                End With
            End If
            If seriescount >= 6 Then
                With cht.SeriesCollection(6)
                    .Border.ColorIndex = xlNone
                    'ForegroundColorIndex lets you set line to no color
                    .MarkerForegroundColorIndex = xlColorIndexNone
                    'BackgroundColor is marker fill
                    .MarkerBackgroundColor = giRGBgreencolor
                End With
            End If

        ElseIf seriescount > 6 Then
            If cht.HasTitle = True Then
                cht.ChartTitle.Text = "You have too many data series for this chart type. Please contact the Communications Department for further guidance."
            ElseIf cht.HasTitle = False Then
                'Add text box with text, size, and use correct font
                Set txtB = cht.TextBoxes.Add(0, 0, 500, 40)
                'Set txtB = cht.TextBoxes.Add(400, 100, 125, 20)
                '(horizontal placement, vertical placement, box width, box height)
                With txtB
                    .name = "TitleBox"
                    .Text = "You have too many data series for this chart type. Please contact the Communications Department for further guidance."
                    .Font.Size = 10
                    .Font.name = "Lato"
                End With
                'Left-align text in text box
                cht.Shapes.Range(Array("TitleBox")).TextEffect.Alignment = msoTextEffectAlignmentLeft
            End If

        End If

'    End If

End Sub

'JAS 2017
'JAS 2023
