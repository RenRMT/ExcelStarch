Attribute VB_Name = "modChartGeneric"

Option Explicit


'INSO Chart style Excel add-in. Modified from UrbanGraphingStyle by Jonathan Schwabish.
'Developed & maintained by INSO_HQ_SIO

'================================
Function ApplyFont(sPrimaryFont As String, sSecondaryFont As String, Optional sDefaultFont As String) As String
    '' Returns the most favored font which is installed
    '' sPrimaryFont is name of font you want to use if it's installed
    '' sSecondaryFont is name of font you would otherwise use
    '' [optional] sDefaultFont is name of font you'd settle for, which should be installed on any system
    Dim bTest As Boolean, sTest As String
    If Len(sDefaultFont) = 0 Then
        sDefaultFont = "Arial" ' a font guaranteed to be present--may default to "MS Sans Serif"
    End If
    sTest = sDefaultFont
    bTest = FontIsInstalled(sTest)
    If bTest Then
        ApplyFont = sDefaultFont
    Else
        ApplyFont = sTest
    End If
    sTest = sSecondaryFont
    bTest = FontIsInstalled(sTest)
    If bTest Then
        ApplyFont = sSecondaryFont
    Else
        ' secondary font is not installed on system
    End If
    sTest = sPrimaryFont
    bTest = FontIsInstalled(sTest)
    If bTest Then
        ApplyFont = sPrimaryFont
    Else
        ' primary font is not installed on system
    End If
End Function

'================================
Function FontIsInstalled(sFont As String) As Boolean
    '' Tools > References > OLE Automation
    Dim NewFont As StdFont
    On Error Resume Next
    ' Create a temporary StdFont object
    Set NewFont = New StdFont
    With NewFont
        ' Assign the proposed font name
        .name = sFont
        ' Return true if font assignment succeded
        FontIsInstalled = (StrComp(sFont, .name, vbTextCompare) = 0)
        ' return actual font name through arguments
        sFont = .name
    End With
    Set NewFont = Nothing
End Function

'================================
Function SetWebVersionRemoveLegend() As VbMsgBoxResult

    Dim answer As VbMsgBoxResult
    answer = MsgBox("This option can only be applied if you have already applied an INSO style to your chart." & vbCrLf & _
                    "Do you want to format your graph for the web (e.g., blog posts)?" & vbNewLine & "(This option adds the Urban Insitute Logo, Source notes, Title and Subtitle)", vbQuestion + vbDefaultButton2 + vbYesNoCancel)

    Select Case answer
    Case vbYes
        gWebVersion = True
    Case vbNo
        gWebVersion = False
    End Select

    SetWebVersionRemoveLegend = answer

End Function


'================================
Function SetPieChartSizeandTitle(cht As Chart) As Boolean

    Dim xaxisB As Shape
    Dim titleB1 As TextBox
    Dim titleB2 As TextBox
    Dim titleB3 As TextBox

    'Variables for centering the pie chart
    Dim chtObj As ChartObject

    Dim chtTop As Double
    Dim chtLeft As Double
    Dim chtWidth As Double
    Dim chtHeight As Double

    Dim pltWidth As Double
    Dim pltHeight As Double
    Dim movechart As Double
    Dim legendWidth As Double

    If gWebVersion Then 'web graphs
        With cht.Parent
            .Width = gdChartWidth_web
            .Height = gdChartHeight_web
        End With
    Else    'print graphs
        With cht.Parent
            .Width = gdChartWidth_print
            .Height = gdChartHeight_print
        End With
    End If

    'Change font for whole chart
    cht.ChartArea.Font.name = FontStyle

    If gWebVersion Then
        'WEB graph title formatting
        If cht.HasTitle = True Then
            cht.ChartTitle.Select
            Selection.Delete
        End If
        Set titleB1 = cht.TextBoxes.Add(0, 0, titleBoxWidth, pieTitleBoxHeight)
        With titleB1
            .name = "TitleBox"
            .Text = "Title in 18pt Title Case"
            .Font.Size = pieTitleFontSize
            .Font.name = FontStyle
            .Font.Bold = msoFalse
        End With
        Set titleB2 = cht.TextBoxes.Add(0, pieSubtitleBoxTop, titleBoxWidth, pieSubtitleBoxHeight)
        With titleB2
            .name = "SubTitleBox"
            .Text = "Subtitle in 14pt sentence case"
            .Font.Size = pieSubtitleFontSize
            .Font.name = FontStyle
            .Font.Bold = msoFalse
        End With
        Set titleB3 = cht.TextBoxes.Add(0, pieYAxisLabelBoxTop, titleBoxWidth, yAxisLabel_noLegendHeight)
        With titleB3
            .name = "YAxisLabelBox"
            .Text = "Y axis title (unit)"
            .Font.Size = axisFontSize
            .Font.name = FontStyleItalic
            .Font.Bold = msoFalse
            .Font.Italic = msoTrue
        End With
    Else
        'PRINT graph title formatting
        If cht.HasTitle = True Then
            cht.ChartTitle.Select
            Selection.Delete
        End If
        Set titleB3 = cht.TextBoxes.Add(0, 0, titleBoxWidth, pieTitleBoxHeight_print)
        With titleB3
            .name = "YAxisLabelBox"
            .Text = "Y axis title (unit)"
            .Font.Size = tickLabelSize_print
            .Font.name = FontStyleItalic
            .Font.Bold = msoFalse
            .Font.Italic = msoTrue
        End With
    End If

    'Size the pie chart
    If gWebVersion Then
        If cht.hasLegend = True Then
            cht.PlotArea.Select
            Selection.Width = piePlotAreaSize_legend
            Selection.Height = piePlotAreaSize_legend
            Selection.Left = piePlotAreaLeft_web
            Selection.Top = piePlotAreaTop_web
        Else
            cht.PlotArea.Select
            Selection.Width = piePlotAreaSize_noLegend
            Selection.Height = piePlotAreaSize_noLegend
            Selection.Left = piePlotAreaLeft_web
            Selection.Top = piePlotAreaTop_web
        End If
    Else
        cht.PlotArea.Select 'this print size works for those with and w/o legend
        Selection.Height = piePlotAreaHeight_print
    End If

    'Position the pie chart (centered horizontally)
    If gWebVersion Then    'web graphs
        Set chtObj = ActiveChart.Parent
        With chtObj
            chtTop = .Chart.ChartArea.Top
            chtLeft = .Chart.ChartArea.Left
            chtHeight = .Chart.ChartArea.Height
            chtWidth = .Chart.ChartArea.Width

            pltHeight = .Chart.PlotArea.Height
            pltWidth = .Chart.PlotArea.Width

            .Chart.PlotArea.Top = (chtHeight - pltHeight) * piePlotTopRatio_web
            .Chart.PlotArea.Left = (chtWidth - pltWidth) / 2
        End With
    Else
        Set chtObj = ActiveChart.Parent
        With chtObj
            chtTop = .Chart.ChartArea.Top
            chtLeft = .Chart.ChartArea.Left
            chtHeight = .Chart.ChartArea.Height
            chtWidth = .Chart.ChartArea.Width

            pltHeight = .Chart.PlotArea.Height
            pltWidth = .Chart.PlotArea.Width

            .Chart.PlotArea.Top = (chtHeight - pltHeight) * piePlotTopRatio_print
            .Chart.PlotArea.Left = (chtWidth - pltWidth) / 2
        End With
    End If

    'Position the legend
    If gWebVersion Then
        If cht.hasLegend = True Then
            ActiveChart.Legend.Select
            Selection.Font.Size = axisFontSize
            'Put legend at top
            cht.Legend.Position = xlLegendPositionTop 'this centers horizontally
            ActiveChart.Legend.Select
            Selection.Top = pieLegendTop_web
        End If
    Else
        If cht.hasLegend = True Then
            ActiveChart.Legend.Select
            Selection.Top = pieLegendTop_print
            Selection.Font.Size = dataLabelFontSize_print
            'Put legend at top
            cht.Legend.Position = xlLegendPositionTop 'this centers horizontally
        End If
    End If

End Function


'================================
Function RemoveLegendandResize(cht As Chart) As Boolean

    If gWebVersion Then
        'Remove Legend
        If ActiveChart.hasLegend = True Then
            ActiveChart.Legend.Select
            Selection.Delete
        End If

        'Set plot area for web graphics
        ActiveChart.PlotArea.Select
        Selection.Height = removeLegend_webHeight
        Selection.Top = removeLegend_webTop
        Selection.Width = removeLegend_webWidth
        Selection.Left = removeLegend_webLeft

    Else

        'Remove Legend
        If ActiveChart.hasLegend = True Then
            ActiveChart.Legend.Select
            Selection.Delete
        End If

        'Set plot area for print graphics
        ActiveChart.PlotArea.Select
        Selection.Height = removeLegend_printHeight
        Selection.Top = removeLegend_printTop
        Selection.Width = removeLegend_printWidth
        Selection.Left = removeLegend_printLeft

    End If

End Function


'================================
Function SlopeChartStyles(cht As Chart) As Boolean
    Dim seriescount As Long
    Dim iseries As Long

    With ActiveChart
        seriescount = .SeriesCollection.Count
    End With

    'Instead of using the FormatXAxis function, do it directly here
    If gWebVersion Then
        With cht.Axes(xlCategory).TickLabels.Font
            .Size = axisFontSize
        End With
    Else
        With cht.Axes(xlCategory).TickLabels.Font
            .Size = tickLabelSize_print
        End With
    End If

    If gWebVersion Then
        For iseries = 1 To seriescount
            ActiveChart.SeriesCollection(iseries).Points(1).DataLabel.Select
            'DoEvents code here forces code to pause;
            'without this, code was crashing in Excel 2016 PC version
            DoEvents
            DoEvents
            DoEvents
            DoEvents
            DoEvents
            DoEvents
            If iseries < 10 Then
                Selection.Format.TextFrame2.TextRange.Characters(1, 7).Font.Bold = msoTrue
                Selection.Format.TextFrame2.TextRange.Characters(1, 7).Font.Size = axisFontSize
                Selection.Format.TextFrame2.TextRange.Characters(8, 2).Font.Bold = msoFalse
                Selection.Format.TextFrame2.TextRange.Characters(8, 2).Font.Size = dataLabelFontSize_secondary
            ElseIf iseries >= 10 Then
                Selection.Format.TextFrame2.TextRange.Characters(1, 8).Font.Bold = msoTrue
                Selection.Format.TextFrame2.TextRange.Characters(1, 8).Font.Size = axisFontSize
                Selection.Format.TextFrame2.TextRange.Characters(10, 2).Font.Bold = msoFalse
                Selection.Format.TextFrame2.TextRange.Characters(10, 2).Font.Size = dataLabelFontSize_secondary
            End If

            ActiveChart.SeriesCollection(iseries).Points(2).DataLabel.Select
            Selection.Format.TextFrame2.TextRange.Font.Size = dataLabelFontSize_secondary
        Next
    Else
        For iseries = 1 To seriescount
            ActiveChart.SeriesCollection(iseries).Points(1).DataLabel.Select
            'DoEvents code here forces code to pause;
            'without this, code was crashing in Excel 2016 PC version
            DoEvents
            DoEvents
            DoEvents
            DoEvents
            DoEvents
            DoEvents
            If iseries < 10 Then
                Selection.Format.TextFrame2.TextRange.Characters(1, 7).Font.Bold = msoTrue
                Selection.Format.TextFrame2.TextRange.Characters(1, 7).Font.Size = axisFontSize
                Selection.Format.TextFrame2.TextRange.Characters(8, 2).Font.Bold = msoFalse
                Selection.Format.TextFrame2.TextRange.Characters(8, 2).Font.Size = dataLabelFontSize_secondary
            ElseIf iseries >= 10 Then
                Selection.Format.TextFrame2.TextRange.Characters(1, 8).Font.Bold = msoTrue
                Selection.Format.TextFrame2.TextRange.Characters(1, 8).Font.Size = axisFontSize
                Selection.Format.TextFrame2.TextRange.Characters(10, 2).Font.Bold = msoFalse
                Selection.Format.TextFrame2.TextRange.Characters(10, 2).Font.Size = dataLabelFontSize_secondary
            End If
            ActiveChart.SeriesCollection(iseries).Points(2).DataLabel.Select
            Selection.Format.TextFrame2.TextRange.Font.Size = tickLabelSize_print
        Next
    End If

    'Squeeze the plot area in from the left to make room for labels
    If gWebVersion Then
        ActiveChart.PlotArea.Select
        Selection.Width = slopePlotWidth_web
        Selection.Left = slopePlotLeft
        Selection.Top = slopePlotTop
        If ActiveChart.hasLegend = True Then
            ActiveChart.Legend.Select
            Selection.Top = legend_top
            Selection.Font.Size = axisFontSize
        End If
    Else
        ActiveChart.PlotArea.Select
        Selection.Width = slopePlotWidth_print
        Selection.Left = slopePlotLeft
        Selection.Top = slopePlotTop
        If ActiveChart.hasLegend = True Then
            ActiveChart.Legend.Select
            Selection.Top = slopeLegendTop_print
            Selection.Font.Size = dataLabelFontSize_print
            Selection.Left = (ActiveChart.ChartArea.Width - ActiveChart.Legend.Width) / 2 + slopeLegend_printLeftPad
        End If
    End If

    'Remove border
    ActiveChart.ChartArea.Border.LineStyle = xlNone

End Function


'================================
Function DotPlotStyles(cht As Chart) As Boolean

    Dim ptscount As Long
    Dim ipts As Long
    Dim srs As Series
    Dim scle As Long

    For Each srs In ActiveChart.SeriesCollection
        With srs
            ptscount = .Points.Count
            If gWebVersion Then
                If ptscount < 9 Then
                    For ipts = 1 To ptscount
                        ActiveChart.SeriesCollection(1).DataLabels.Select
                        ActiveChart.SeriesCollection(1).Points(ipts).DataLabel.Select
                        Selection.Format.TextFrame2.TextRange.Font.Size = axisFontSize
                        Selection.Format.TextFrame2.TextRange.Characters(1, 7).Font.Bold = msoTrue
                        Selection.Format.TextFrame2.TextRange.Characters(8, 3).Font.Size = dataLabelFontSize_secondary
                        ActiveChart.SeriesCollection(2).DataLabels.Select
                        Selection.Format.TextFrame2.TextRange.Font.Size = dataLabelFontSize_secondary
                    Next
                ElseIf ptscount = 9 Then
                    For ipts = 1 To 8
                        ActiveChart.SeriesCollection(1).DataLabels.Select
                        ActiveChart.SeriesCollection(1).Points(ipts).DataLabel.Select
                        Selection.Format.TextFrame2.TextRange.Font.Size = axisFontSize
                        Selection.Format.TextFrame2.TextRange.Characters(1, 7).Font.Bold = msoTrue
                        Selection.Format.TextFrame2.TextRange.Characters(8, 3).Font.Size = dataLabelFontSize_secondary
                        ActiveChart.SeriesCollection(2).DataLabels.Select
                        Selection.Format.TextFrame2.TextRange.Font.Size = dataLabelFontSize_secondary
                    Next
                    For ipts = 9 To 9
                        ActiveChart.SeriesCollection(1).DataLabels.Select
                        ActiveChart.SeriesCollection(1).Points(ipts).DataLabel.Select
                        Selection.Format.TextFrame2.TextRange.Font.Size = axisFontSize
                        Selection.Format.TextFrame2.TextRange.Characters(1, 7).Font.Bold = msoTrue
                        Selection.Format.TextFrame2.TextRange.Characters(8, 4).Font.Size = dataLabelFontSize_secondary
                        ActiveChart.SeriesCollection(2).DataLabels.Select
                      Selection.Format.TextFrame2.TextRange.Font.Size = dataLabelFontSize_secondary
                    Next
                ElseIf ptscount > 9 Then
                    For ipts = 1 To 8
                        ActiveChart.SeriesCollection(1).DataLabels.Select
                        ActiveChart.SeriesCollection(1).Points(ipts).DataLabel.Select
                        Selection.Format.TextFrame2.TextRange.Font.Size = axisFontSize
                        Selection.Format.TextFrame2.TextRange.Characters(1, 7).Font.Bold = msoTrue
                        Selection.Format.TextFrame2.TextRange.Characters(8, 3).Font.Size = dataLabelFontSize_secondary
                        ActiveChart.SeriesCollection(2).DataLabels.Select
                        Selection.Format.TextFrame2.TextRange.Font.Size = dataLabelFontSize_secondary
                    Next
                    For ipts = 9 To 9
                        ActiveChart.SeriesCollection(1).DataLabels.Select
                        ActiveChart.SeriesCollection(1).Points(ipts).DataLabel.Select
                        Selection.Format.TextFrame2.TextRange.Font.Size = axisFontSize
                        Selection.Format.TextFrame2.TextRange.Characters(1, 7).Font.Bold = msoTrue
                        Selection.Format.TextFrame2.TextRange.Characters(8, 4).Font.Size = dataLabelFontSize_secondary
                        ActiveChart.SeriesCollection(2).DataLabels.Select
                      Selection.Format.TextFrame2.TextRange.Font.Size = dataLabelFontSize_secondary
                    Next
                    For ipts = 10 To ptscount
                        ActiveChart.SeriesCollection(1).DataLabels.Select
                        ActiveChart.SeriesCollection(1).Points(ipts).DataLabel.Select
                        Selection.Format.TextFrame2.TextRange.Font.Size = axisFontSize
                        Selection.Format.TextFrame2.TextRange.Characters(1, 8).Font.Bold = msoTrue
                        Selection.Format.TextFrame2.TextRange.Characters(9, 4).Font.Size = dataLabelFontSize_secondary
                        ActiveChart.SeriesCollection(2).DataLabels.Select
                        Selection.Format.TextFrame2.TextRange.Font.Size = dataLabelFontSize_secondary
                    Next
                End If
            Else 'print
                If ptscount < 9 Then
                    For ipts = 1 To ptscount
                        ActiveChart.SeriesCollection(1).DataLabels.Select
                        ActiveChart.SeriesCollection(1).Points(ipts).DataLabel.Select
                        Selection.Format.TextFrame2.TextRange.Font.Size = dataLabelFontSize_print
                        Selection.Format.TextFrame2.TextRange.Characters(1, 7).Font.Bold = msoTrue
                        Selection.Format.TextFrame2.TextRange.Characters(8, 3).Font.Size = tickLabelSize_print
                        ActiveChart.SeriesCollection(2).DataLabels.Select
                        Selection.Format.TextFrame2.TextRange.Font.Size = tickLabelSize_print
                    Next
                ElseIf ptscount = 9 Then
                    For ipts = 1 To 8
                        ActiveChart.SeriesCollection(1).DataLabels.Select
                        ActiveChart.SeriesCollection(1).Points(ipts).DataLabel.Select
                        Selection.Format.TextFrame2.TextRange.Font.Size = dataLabelFontSize_print
                        Selection.Format.TextFrame2.TextRange.Characters(1, 7).Font.Bold = msoTrue
                        Selection.Format.TextFrame2.TextRange.Characters(8, 3).Font.Size = tickLabelSize_print
                        ActiveChart.SeriesCollection(2).DataLabels.Select
                        Selection.Format.TextFrame2.TextRange.Font.Size = tickLabelSize_print
                    Next
                    For ipts = 9 To 9
                        ActiveChart.SeriesCollection(1).DataLabels.Select
                        ActiveChart.SeriesCollection(1).Points(ipts).DataLabel.Select
                        Selection.Format.TextFrame2.TextRange.Font.Size = dataLabelFontSize_print
                        Selection.Format.TextFrame2.TextRange.Characters(1, 7).Font.Bold = msoTrue
                        Selection.Format.TextFrame2.TextRange.Characters(8, 4).Font.Size = tickLabelSize_print
                        ActiveChart.SeriesCollection(2).DataLabels.Select
                        Selection.Format.TextFrame2.TextRange.Font.Size = tickLabelSize_print
                    Next
                ElseIf ptscount > 9 Then
                    For ipts = 1 To 8
                        ActiveChart.SeriesCollection(1).DataLabels.Select
                        ActiveChart.SeriesCollection(1).Points(ipts).DataLabel.Select
                        Selection.Format.TextFrame2.TextRange.Font.Size = dataLabelFontSize_print
                        Selection.Format.TextFrame2.TextRange.Characters(1, 7).Font.Bold = msoTrue
                        Selection.Format.TextFrame2.TextRange.Characters(8, 3).Font.Size = tickLabelSize_print
                        ActiveChart.SeriesCollection(2).DataLabels.Select
                        Selection.Format.TextFrame2.TextRange.Font.Size = tickLabelSize_print
                    Next
                    For ipts = 9 To 9
                        ActiveChart.SeriesCollection(1).DataLabels.Select
                        ActiveChart.SeriesCollection(1).Points(ipts).DataLabel.Select
                        Selection.Format.TextFrame2.TextRange.Font.Size = dataLabelFontSize_print
                        Selection.Format.TextFrame2.TextRange.Characters(1, 7).Font.Bold = msoTrue
                        Selection.Format.TextFrame2.TextRange.Characters(8, 4).Font.Size = tickLabelSize_print
                        ActiveChart.SeriesCollection(2).DataLabels.Select
                        Selection.Format.TextFrame2.TextRange.Font.Size = tickLabelSize_print
                    Next
                    For ipts = 10 To ptscount
                        ActiveChart.SeriesCollection(1).DataLabels.Select
                        ActiveChart.SeriesCollection(1).Points(ipts).DataLabel.Select
                        Selection.Format.TextFrame2.TextRange.Font.Size = dataLabelFontSize_print
                        Selection.Format.TextFrame2.TextRange.Characters(1, 8).Font.Bold = msoTrue
                        Selection.Format.TextFrame2.TextRange.Characters(9, 4).Font.Size = tickLabelSize_print
                        ActiveChart.SeriesCollection(2).DataLabels.Select
                        Selection.Format.TextFrame2.TextRange.Font.Size = tickLabelSize_print
                    Next
                End If
            End If
        End With
    Next

End Function


Function ScatterplotStyles(cht As Chart) As Boolean

    'Format vertical gridlines
    If gWebVersion Then
        If cht.Axes(xlCategory).HasMajorGridlines = True Then
            cht.Axes(xlCategory).MajorGridlines.Select
            With Selection.Format.Line
                .Visible = msoTrue
                .weight = gridlineWeight
                .DashStyle = msoLineSysDot
                .ForeColor.rgb = giRGBgridlinesweb
            End With
        Else    'if no vertical gridlines, then add them
            cht.SetElement (msoElementPrimaryCategoryGridLinesMajor)
            cht.Axes(xlCategory).MajorGridlines.Select
            With Selection.Format.Line
                .Visible = msoTrue
                .weight = gridlineWeight
                .DashStyle = msoLineSysDot
                .ForeColor.rgb = giRGBgridlinesweb
            End With
        End If
    Else
        If cht.Axes(xlCategory).HasMajorGridlines = True Then
            cht.Axes(xlCategory).MajorGridlines.Select
            With Selection.Format.Line
                .Visible = msoTrue
                .weight = gridlineWeight
                .DashStyle = msoLineSysDot
                .ForeColor.rgb = giRGBgridlinesprint
            End With
        Else    'if no vertical gridlines, then add them
            cht.SetElement (msoElementPrimaryCategoryGridLinesMajor)
            cht.Axes(xlCategory).MajorGridlines.Select
            With Selection.Format.Line
                .Visible = msoTrue
                .weight = gridlineWeight
                .DashStyle = msoLineSysDot
                .ForeColor.rgb = giRGBgridlinesprint
            End With
        End If
    End If    'web/print for vertical lines

    'Format horizontal gridlines
    If gWebVersion Then
        If cht.Axes(xlValue).HasMajorGridlines = True Then
            cht.Axes(xlValue).MajorGridlines.Select
            With Selection.Format.Line
                .Visible = msoTrue
                .weight = gridlineWeight
                .DashStyle = msoLineSysDot
                .ForeColor.rgb = giRGBgridlinesweb
            End With
        Else    'add gridlines if they don't exist
            cht.PlotArea.Select
            cht.SetElement (msoElementPrimaryValueGridLinesMajor)
            cht.Axes(xlValue).MajorGridlines.Select
            With Selection.Format.Line
                .Visible = msoTrue
                .weight = gridlineWeight
                .DashStyle = msoLineSysDot
                .ForeColor.rgb = giRGBgridlinesweb
            End With
        End If
    Else
        If cht.Axes(xlValue).HasMajorGridlines = True Then
            cht.Axes(xlValue).MajorGridlines.Select
            With Selection.Format.Line
                .Visible = msoTrue
                .weight = gridlineWeight
                .DashStyle = msoLineSysDot
                .ForeColor.rgb = giRGBgridlinesprint
            End With
        Else    'add gridlines if they don't exist
            cht.PlotArea.Select
            cht.SetElement (msoElementPrimaryValueGridLinesMajor)
            cht.Axes(xlValue).MajorGridlines.Select
            With Selection.Format.Line
                .Visible = msoTrue
                .weight = gridlineWeight
                .DashStyle = msoLineSysDot
                .ForeColor.rgb = giRGBgridlinesprint
            End With
        End If
    End If    'web/print for horizontal lines

End Function
