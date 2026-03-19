Attribute VB_Name = "m_UrbanLabelLastPoint"
Option Explicit



Sub LabelLastPointButton()
    Dim ipts As Long
    Dim Npts As Long
    Dim bLabeled As Boolean
    Dim mySrs As Series
    Dim cht As Chart
    Dim srs As Series
    Dim bRemoveLegendandResize As Boolean
    Dim plHeight As Double
    Dim plWidth As Long
    Dim shp As Shape
    Dim iColor As Long
    
    'Check ActiveChart
    If ActiveChart Is Nothing Then
        MsgBox "First, create a chart, or select an existing chart, and try again.", vbExclamation, "No Active Chart"
    Else

        'Duplicate the selected chart
        ActiveChart.Parent.Duplicate.Select
        ActiveChart.PlotArea.Select
        Selection.Width = gdChartWidth_web - 50
        Selection.Left = 0

        Set cht = ActiveChart

        'Adjust size and placement of title
        If gWebVersion Then
            If cht.hasLegend = True Then
                For Each shp In cht.Shapes
                    If shp.name = "YAxisLabelBox" Then
                        cht.Shapes.Range(Array("YAxisLabelBox")).Select
                        Selection.ShapeRange.IncrementTop -20
                    End If
                Next
            Else
                For Each shp In cht.Shapes
                    If shp.name = "YAxisLabelBox" Then
                        cht.Shapes.Range(Array("YAxisLabelBox")).Select
                        'Selection.ShapeRange.IncrementTop -20
                    End If
                Next
            End If 'HasLegend
            With cht.PlotArea.Select
                plHeight = cht.PlotArea.Height
                plWidth = cht.PlotArea.Width
                cht.PlotArea.Select
                'Selection.Height = plHeight * 2 '1.25
                Selection.Top = 60
                Selection.Width = plWidth * 0.98
                If cht.hasLegend = True Then
                    Selection.Height = plHeight * 1.15
                Else
                    Selection.Height = plHeight * 1
                End If
            End With
        Else
        End If

        'Remove Legend
        'Note: this does not call the CommonFunctions version so that it can be applied to any chart)
        If cht.hasLegend = True Then
            ActiveChart.Legend.Select
            Selection.Delete
        End If

        'For Each mySrs In ActiveChart.SeriesCollection
        For Each srs In ActiveChart.SeriesCollection

            bLabeled = False
            With srs
                Npts = 0
                On Error Resume Next
                Npts = .Points.Count
                On Error GoTo 0
                If Npts > 0 Then
                    ' start at last point, work backwards
                    ' label last valid point, remove labels on earlier points
                    For ipts = Npts To 1 Step -1
                        If bLabeled Then
                            ' handle error if point isn't plotted
                            On Error Resume Next
                            ' remove existing label if it's not the last point
                            srs.Points(ipts).HasDataLabel = False
                            On Error GoTo 0
                        Else
                            ' handle error if point isn't plotted
                            On Error Resume Next
                            ' remove existing label (linked labels otherwise resist reassignment)
                            srs.Points(ipts).HasDataLabel = False
                            ' add label
                            srs.Points(ipts).ApplyDataLabels _
                                    ShowSeriesName:=True, ShowCategoryName:=False, _
                                    ShowValue:=False, AutoText:=False, LegendKey:=False
                            ' 2003 error trying to label unplotted point
                            bLabeled = (Err.Number = 0)
                            ' 2010 no error if point doesn't exist, label applied, but it's blank
                            If bLabeled Then bLabeled = (Len(srs.Points(ipts).DataLabel.Text) > 0)
                            If Not bLabeled Then
                                srs.Points(ipts).HasDataLabel = False
                            End If
                            If bLabeled Then
                                ' data label position
                                Select Case srs.chartType
                                Case xlLine, xlLineStacked, xlLineStacked100, xlLineMarkers, xlLineMarkersStacked, xlLineMarkersStacked100
                                    srs.Points(ipts).DataLabel.Position = xlLabelPositionRight
                                    srs.Points(ipts).DataLabel.Font.Bold = msoTrue
                                    iColor = .Format.Line.ForeColor.rgb
                                    If gWebVersion Then    'web graphs
                                        srs.Points(ipts).DataLabel.Font.Size = 12
                                        srs.Points(ipts).DataLabel.Font.color = iColor
                                    Else
                                        srs.Points(ipts).DataLabel.Font.Size = 9.5
                                        srs.Points(ipts).DataLabel.Font.color = iColor
                                    End If
                                Case xlXYScatter, xlXYScatterLines, xlXYScatterLinesNoMarkers, xlXYScatterSmooth, xlXYScatterSmoothNoMarkers
                                    srs.Points(ipts).DataLabel.Position = xlLabelPositionRight
                                    srs.Points(ipts).DataLabel.Font.Bold = msoTrue
                                    iColor = .MarkerBackgroundColor
                                    If gWebVersion Then    'web graphs
                                        srs.Points(ipts).DataLabel.Font.Size = 12
                                        srs.Points(ipts).DataLabel.Font.color = iColor
                                    Else
                                        srs.Points(ipts).DataLabel.Font.Size = 9.5
                                        srs.Points(ipts).DataLabel.Font.color = iColor
                                    End If
                                Case xlColumnClustered, xlBarClustered
                                    srs.Points(ipts).DataLabel.Position = xlLabelPositionOutsideEnd
                                    srs.Points(ipts).DataLabel.Font.Bold = msoTrue
                                    iColor = .Format.Fill.ForeColor.rgb
                                    If gWebVersion Then    'web graphs
                                        srs.Points(ipts).DataLabel.Font.Size = 12
                                        srs.Points(ipts).DataLabel.Font.color = iColor
                                    Else
                                        srs.Points(ipts).DataLabel.Font.Size = 9.5
                                        srs.Points(ipts).DataLabel.Font.color = iColor
                                    End If
                                Case xlColumnStacked, xlColumnStacked100, xlBarStacked, xlBarStacked100, xlArea, xlAreaStacked, xlAreaStacked100
                                    srs.Points(ipts).DataLabel.Position = xlLabelPositionCenter
                                    srs.Points(ipts).DataLabel.Font.Bold = msoTrue
                                    iColor = .Format.Fill.ForeColor.rgb
                                    If gWebVersion Then    'web graphs
                                        srs.Points(ipts).DataLabel.Font.Size = 12
                                        srs.Points(ipts).DataLabel.Font.color = iColor
                                    Else
                                        srs.Points(ipts).DataLabel.Font.Size = 9.5
                                        srs.Points(ipts).DataLabel.Font.color = iColor
                                    End If
                                End Select
                            End If
                            On Error GoTo 0
                        End If
                    Next
                End If
                ' if you don't do this, it won't update if series name changes
                srs.DataLabels.AutoText = True
            End With

        Next

    End If

End Sub



Public Sub LabelLastPointButton_onAction(control As IRibbonControl)
'
' Code for onAction callback. Ribbon control button
'
    Dim ipts As Long
    Dim Npts As Long
    Dim bLabeled As Boolean
    Dim mySrs As Series
    Dim cht As Chart
    Dim srs As Series
    Dim bRemoveLegendandResize As Boolean
    Dim plHeight As Double
    Dim plWidth As Long
    Dim shp As Shape
    Dim iColor As Long
    
    'Check ActiveChart
    If ActiveChart Is Nothing Then
        MsgBox "First, create a chart, or select an existing chart, and try again.", vbExclamation, "No Active Chart"
    Else
    
'''        'No longer need this code. Excel will hold onto the web/print declaration
'''        'from the initial menu and apply it here.
'''''        'SetWebVersion
'''''        'Instead of passing gWebVersion back and forth, use it locally here
'''''        Dim answer As VbMsgBoxResult
'''''        answer = MsgBox("Do you want to format your graph for the WEB (e.g., blog posts)?" & vbNewLine & "(Web version will set labels at 12pt font; otherwise at 9.5pt font)", vbQuestion + vbDefaultButton2 + vbYesNoCancel)
'''''
'''''        Select Case answer
'''''        Case vbYes
'''''            gWebVersion = True
'''''        Case vbNo
'''''            gWebVersion = False
'''''        Case vbCancel
'''''            Exit Sub
'''''        End Select

        'Duplicate the selected chart
        ActiveChart.Parent.Duplicate.Select
        ActiveChart.PlotArea.Select
        'Only need to change width of plot area for a line chart
        If ActiveChart.chartType = xlLine Then
            Selection.Width = gdChartWidth_web - 50
        End If
        Selection.Left = 0

        Set cht = ActiveChart

        'Adjust size and placement of title
       'If gWebVersion Then
            If cht.hasLegend = True Then
                For Each shp In cht.Shapes
                    If shp.name = "YAxisLabelBox" Then
                        cht.Shapes.Range(Array("YAxisLabelBox")).Select
                        Selection.ShapeRange.IncrementTop -10
                    End If
                Next
            Else
                For Each shp In cht.Shapes
                    If shp.name = "YAxisLabelBox" Then
                        cht.Shapes.Range(Array("YAxisLabelBox")).Select
                        'Selection.ShapeRange.IncrementTop -20
                    End If
                Next
            End If 'HasLegend
            With cht.PlotArea.Select
                plHeight = cht.PlotArea.Height
                plWidth = cht.PlotArea.Width
                cht.PlotArea.Select
                'Selection.Height = plHeight * 2 '1.25
                Selection.Top = 80
                If gWebVersion Then
                    Selection.Width = plWidth * 0.98
                Else
                    Selection.Width = plWidth * 0.9
                End If
                If cht.hasLegend = True Then
                    Selection.Height = plHeight * 1
                Else
                    Selection.Height = plHeight * 1
                End If
            End With
        'Else
        'End If

        'Remove Legend
        'Note: this does not call the CommonFunctions version so that it can be applied to any chart)
        If cht.hasLegend = True Then
            ActiveChart.Legend.Select
            Selection.Delete
        End If

        'For Each mySrs In ActiveChart.SeriesCollection
        For Each srs In ActiveChart.SeriesCollection

            bLabeled = False
            With srs
                Npts = 0
                On Error Resume Next
                Npts = .Points.Count
                On Error GoTo 0
                If Npts > 0 Then
                    ' start at last point, work backwards
                    ' label last valid point, remove labels on earlier points
                    For ipts = Npts To 1 Step -1
                        If bLabeled Then
                            ' handle error if point isn't plotted
                            On Error Resume Next
                            ' remove existing label if it's not the last point
                            srs.Points(ipts).HasDataLabel = False
                            On Error GoTo 0
                        Else
                            ' handle error if point isn't plotted
                            On Error Resume Next
                            ' remove existing label (linked labels otherwise resist reassignment)
                            srs.Points(ipts).HasDataLabel = False
                            ' add label
                            srs.Points(ipts).ApplyDataLabels _
                                    ShowSeriesName:=True, ShowCategoryName:=False, _
                                    ShowValue:=False, AutoText:=False, LegendKey:=False
                            ' 2003 error trying to label unplotted point
                            bLabeled = (Err.Number = 0)
                            ' 2010 no error if point doesn't exist, label applied, but it's blank
                            If bLabeled Then bLabeled = (Len(srs.Points(ipts).DataLabel.Text) > 0)
                            If Not bLabeled Then
                                srs.Points(ipts).HasDataLabel = False
                            End If
                            If bLabeled Then
                                ' data label position
                                Select Case srs.chartType
                                Case xlLine, xlLineStacked, xlLineStacked100, xlLineMarkers, xlLineMarkersStacked, xlLineMarkersStacked100
                                    srs.Points(ipts).DataLabel.Position = xlLabelPositionRight
                                    srs.Points(ipts).DataLabel.Font.Bold = msoTrue
                                    iColor = .Format.Line.ForeColor.rgb
                                    If gWebVersion Then    'web graphs
                                        srs.Points(ipts).DataLabel.Font.Size = 12
                                        srs.Points(ipts).DataLabel.Font.color = iColor
                                    Else
                                        srs.Points(ipts).DataLabel.Font.Size = 9.5
                                        srs.Points(ipts).DataLabel.Font.color = iColor
                                    End If
                                Case xlXYScatter, xlXYScatterLines, xlXYScatterLinesNoMarkers, xlXYScatterSmooth, xlXYScatterSmoothNoMarkers
                                    srs.Points(ipts).DataLabel.Position = xlLabelPositionRight
                                    srs.Points(ipts).DataLabel.Font.Bold = msoTrue
                                    iColor = .MarkerBackgroundColor
                                    If gWebVersion Then    'web graphs
                                        srs.Points(ipts).DataLabel.Font.Size = 12
                                        srs.Points(ipts).DataLabel.Font.color = iColor
                                    Else
                                        srs.Points(ipts).DataLabel.Font.Size = 9.5
                                        srs.Points(ipts).DataLabel.Font.color = iColor
                                    End If
                                Case xlColumnClustered, xlBarClustered
                                    srs.Points(ipts).DataLabel.Position = xlLabelPositionOutsideEnd
                                    srs.Points(ipts).DataLabel.Font.Bold = msoTrue
                                    iColor = .Format.Fill.ForeColor.rgb
                                    If gWebVersion Then    'web graphs
                                        srs.Points(ipts).DataLabel.Font.Size = 12
                                        srs.Points(ipts).DataLabel.Font.color = iColor
                                    Else
                                        srs.Points(ipts).DataLabel.Font.Size = 9.5
                                        srs.Points(ipts).DataLabel.Font.color = iColor
                                    End If
                                Case xlColumnStacked, xlColumnStacked100, xlBarStacked, xlBarStacked100, xlArea, xlAreaStacked, xlAreaStacked100
                                    srs.Points(ipts).DataLabel.Position = xlLabelPositionCenter
                                    srs.Points(ipts).DataLabel.Font.Bold = msoTrue
                                    iColor = .Format.Fill.ForeColor.rgb
                                    If gWebVersion Then    'web graphs
                                        srs.Points(ipts).DataLabel.Font.Size = 12
                                        srs.Points(ipts).DataLabel.Font.color = iColor
                                    Else
                                        srs.Points(ipts).DataLabel.Font.Size = 9.5
                                        srs.Points(ipts).DataLabel.Font.color = iColor
                                    End If
                                End Select
                            End If
                            On Error GoTo 0
                        End If
                    Next
                End If
                ' if you don't do this, it won't update if series name changes
                srs.DataLabels.AutoText = True
            End With

        Next

    End If

End Sub

'JAS 2023
