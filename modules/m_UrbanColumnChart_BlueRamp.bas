Attribute VB_Name = "m_UrbanColumnChart_BlueRamp"
Option Explicit



Sub UrbanColumnChart_BlueRamp()
'
    Dim seriescount As Long
    Dim ishadow As Long
    Dim cht As Chart
    Dim txtB As Shape    ''''Not TextBox
    Dim bInsertLogo As Boolean
    Dim bInsertSource As Boolean
    Dim bFormatTitle As Boolean
    Dim bFormatGridlines As Boolean
    Dim bFormatXAxis As Boolean
    Dim bFormatXAxisTitle As Boolean
    Dim bFormatBlueFillColors As Boolean
    Dim bOuterFormat As Boolean
    Dim bSetWebVersion As Boolean

'    'Check ActiveChart
'    If ActiveChart Is Nothing Then
'        MsgBox "First, create a chart, or select an existing chart, and try again.", vbExclamation, "No Active Chart"
'    Else

        ''''bSetWebVersion = SetWebVersion(gWebCancel)
        SetWebVersion_NEW
        If gWebCancel Then
            Exit Sub
        End If

        ActiveSheet.Shapes.AddChart2(-1, xlColumnClustered).Select
        'Duplicate the selected chart
        ActiveChart.Parent.Duplicate.Select

        Set cht = ActiveChart

        bOuterFormat = OuterFormat(cht)
        bFormatXAxisTitle = FormatXAxisTitle(cht)
        bInsertLogo = InsertLogo(cht)
        bInsertSource = InsertSource(cht)
        bFormatTitle = FormatTitle(cht)
        bFormatGridlines = FormatGridlines(cht)
        bFormatXAxis = FormatXAxis(cht)
        bFormatBlueFillColors = FormatBlueFillColors(cht)
        'Column chart in different order than bar chart

        'Style xaxis tick marks
        cht.Axes(xlCategory).MajorTickMark = xlTickMarkNone
        cht.Axes(xlCategory).MinorTickMark = xlTickMarkNone

        'Set series overlap to 0% and gap width to 70%
        cht.ChartGroups(1).Overlap = 0
        cht.ChartGroups(1).GapWidth = 70

        'Change colors of columns
        With cht
            seriescount = .SeriesCollection.Count
        End With

        'Remove shadow from bars (this is default behavior on Macs)
        For ishadow = 1 To seriescount
            cht.SeriesCollection(ishadow).Select
            With Selection.Format.Shadow
                .Visible = msoFalse
            End With
        Next

        'Change colors of columns
        If seriescount = 1 Then
            cht.SeriesCollection(1).Select
            With Selection.Format.Fill
                .Visible = msoTrue
                .ForeColor.rgb = giRGBbluecolor5
                .Solid
            End With

        ElseIf seriescount = 2 Then
            cht.SeriesCollection(1).Select
            With Selection.Format.Fill
                .Visible = msoTrue
                .ForeColor.rgb = giRGBbluecolor2
                .Solid
            End With
            cht.SeriesCollection(2).Select
            With Selection.Format.Fill
                .Visible = msoTrue
                .ForeColor.rgb = giRGBbluecolor5
                .Solid
            End With

        ElseIf seriescount = 3 Then
            cht.SeriesCollection(1).Select
            With Selection.Format.Fill
                .Visible = msoTrue
                .ForeColor.rgb = giRGBbluecolor2
                .Solid
            End With
            cht.SeriesCollection(2).Select
            With Selection.Format.Fill
                .Visible = msoTrue
                .ForeColor.rgb = giRGBbluecolor5
                .Solid
            End With
            cht.SeriesCollection(3).Select
            With Selection.Format.Fill
                .Visible = msoTrue
                .ForeColor.rgb = giRGBbluecolor7
                .Solid
            End With

        ElseIf seriescount = 4 Then
            cht.SeriesCollection(1).Select
            With Selection.Format.Fill
                .Visible = msoTrue
                .ForeColor.rgb = giRGBbluecolor1
                .Solid
            End With
            cht.SeriesCollection(2).Select
            With Selection.Format.Fill
                .Visible = msoTrue
                .ForeColor.rgb = giRGBbluecolor3
                .Solid
            End With
            cht.SeriesCollection(3).Select
            With Selection.Format.Fill
                .Visible = msoTrue
                .ForeColor.rgb = giRGBbluecolor5
                .Solid
            End With
            cht.SeriesCollection(4).Select
            With Selection.Format.Fill
                .Visible = msoTrue
                .ForeColor.rgb = giRGBbluecolor7
                .Solid
            End With

        ElseIf seriescount = 5 Then
            cht.SeriesCollection(1).Select
            With Selection.Format.Fill
                .Visible = msoTrue
                .ForeColor.rgb = giRGBbluecolor1
                .Solid
            End With
            cht.SeriesCollection(2).Select
            With Selection.Format.Fill
                .Visible = msoTrue
                .ForeColor.rgb = giRGBbluecolor3
                .Solid
            End With
            cht.SeriesCollection(3).Select
            With Selection.Format.Fill
                .Visible = msoTrue
                .ForeColor.rgb = giRGBbluecolor5
                .Solid
            End With
            cht.SeriesCollection(4).Select
            With Selection.Format.Fill
                .Visible = msoTrue
                .ForeColor.rgb = giRGBbluecolor7
                .Solid
            End With
            cht.SeriesCollection(5).Select
            With Selection.Format.Fill
                .Visible = msoTrue
                .ForeColor.rgb = colorBlack
                .Solid
            End With

        ElseIf seriescount = 6 Then
            cht.SeriesCollection(1).Select
            With Selection.Format.Fill
                .Visible = msoTrue
                .ForeColor.rgb = giRGBbluecolor1
                .Solid
            End With
            cht.SeriesCollection(2).Select
            With Selection.Format.Fill
                .Visible = msoTrue
                .ForeColor.rgb = giRGBbluecolor2
                .Solid
            End With
            cht.SeriesCollection(3).Select
            With Selection.Format.Fill
                .Visible = msoTrue
                .ForeColor.rgb = giRGBbluecolor3
                .Solid
            End With
            cht.SeriesCollection(4).Select
            With Selection.Format.Fill
                .Visible = msoTrue
                .ForeColor.rgb = giRGBbluecolor4
                .Solid
            End With
            cht.SeriesCollection(5).Select
            With Selection.Format.Fill
                .Visible = msoTrue
                .ForeColor.rgb = giRGBbluecolor5
                .Solid
            End With
            cht.SeriesCollection(6).Select
            With Selection.Format.Fill
                .Visible = msoTrue
                .ForeColor.rgb = giRGBbluecolor6
                .Solid
            End With

        ElseIf seriescount > 6 Then
            If cht.HasTitle = True Then
                cht.ChartTitle.Text = "You have too many data series for this chart type. Please contact the Communications Department for further guidance."
            ElseIf cht.HasTitle = False Then
                'Add text box with text, size, and use correct font
                ''''Set txtB = cht.TextBoxes.Add(0, 0, 500, 40) '''' JP 2016 12 28 Shape not textbox
                Set txtB = cht.Shapes.AddTextbox(msoTextOrientationHorizontal, 0, 0, 500, 40)
                'Set txtB = cht.TextBoxes.Add(400, 100, 125, 20)
                '(horizontal placement, vertical placement, box width, box height)
                With txtB
                    .name = "TitleBox"
                    With .TextFrame2.TextRange
                        .Text = "You have too many data series for this chart type. Please contact the Communications Department for further guidance."
                        .Font.Size = 10
                        .Font.name = "Lato"
                        .Font.Fill.ForeColor.rgb = vbRed
                        'Left-align text in text box
                        .ParagraphFormat.Alignment = msoTextEffectAlignmentLeft
                    End With
                    .Fill.ForeColor.rgb = vbYellow
                End With
            End If

        End If

'    End If

End Sub

Public Sub ColumnwithBlueRamp_onAction(control As IRibbonControl)
'
' Code for onAction callback. Ribbon control button
'
    Dim seriescount As Long
    Dim ishadow As Long
    Dim cht As Chart
    Dim txtB As Shape    ''''Not TextBox
    Dim bInsertLogo As Boolean
    Dim bInsertSource As Boolean
    Dim bFormatTitle As Boolean
    Dim bFormatGridlines As Boolean
    Dim bFormatXAxis As Boolean
    Dim bFormatXAxisTitle As Boolean
    Dim bFormatBlueFillColors As Boolean
    Dim bOuterFormat As Boolean
    Dim bSetWebVersion As Boolean

'    'Check ActiveChart
'    If ActiveChart Is Nothing Then
'        MsgBox "First, create a chart, or select an existing chart, and try again.", vbExclamation, "No Active Chart"
'    Else

        ''''bSetWebVersion = SetWebVersion(gWebCancel)
        SetWebVersion_NEW
        If gWebCancel Then
            Exit Sub
        End If

        ActiveSheet.Shapes.AddChart2(-1, xlColumnClustered).Select
        'Duplicate the selected chart
        ActiveChart.Parent.Duplicate.Select

        Set cht = ActiveChart

        bOuterFormat = OuterFormat(cht)
        bFormatXAxisTitle = FormatXAxisTitle(cht)
        bInsertLogo = InsertLogo(cht)
        bInsertSource = InsertSource(cht)
        bFormatTitle = FormatTitle(cht)
        bFormatGridlines = FormatGridlines(cht)
        bFormatXAxis = FormatXAxis(cht)
        bFormatBlueFillColors = FormatBlueFillColors(cht)
        'Column chart in different order than bar chart

        'Style xaxis tick marks
        cht.Axes(xlCategory).MajorTickMark = xlTickMarkNone
        cht.Axes(xlCategory).MinorTickMark = xlTickMarkNone

        'Set series overlap to 0% and gap width to 70%
        cht.ChartGroups(1).Overlap = 0
        cht.ChartGroups(1).GapWidth = 70

        'Change colors of columns
        With cht
            seriescount = .SeriesCollection.Count
        End With

        'Remove shadow from bars (this is default behavior on Macs)
        For ishadow = 1 To seriescount
            cht.SeriesCollection(ishadow).Select
            With Selection.Format.Shadow
                .Visible = msoFalse
            End With
        Next

        'Change colors of columns
        If seriescount = 1 Then
            cht.SeriesCollection(1).Select
            With Selection.Format.Fill
                .Visible = msoTrue
                .ForeColor.rgb = giRGBbluecolor5
                .Solid
            End With

        ElseIf seriescount = 2 Then
            cht.SeriesCollection(1).Select
            With Selection.Format.Fill
                .Visible = msoTrue
                .ForeColor.rgb = giRGBbluecolor2
                .Solid
            End With
            cht.SeriesCollection(2).Select
            With Selection.Format.Fill
                .Visible = msoTrue
                .ForeColor.rgb = giRGBbluecolor5
                .Solid
            End With

        ElseIf seriescount = 3 Then
            cht.SeriesCollection(1).Select
            With Selection.Format.Fill
                .Visible = msoTrue
                .ForeColor.rgb = giRGBbluecolor2
                .Solid
            End With
            cht.SeriesCollection(2).Select
            With Selection.Format.Fill
                .Visible = msoTrue
                .ForeColor.rgb = giRGBbluecolor5
                .Solid
            End With
            cht.SeriesCollection(3).Select
            With Selection.Format.Fill
                .Visible = msoTrue
                .ForeColor.rgb = giRGBbluecolor7
                .Solid
            End With

        ElseIf seriescount = 4 Then
            cht.SeriesCollection(1).Select
            With Selection.Format.Fill
                .Visible = msoTrue
                .ForeColor.rgb = giRGBbluecolor1
                .Solid
            End With
            cht.SeriesCollection(2).Select
            With Selection.Format.Fill
                .Visible = msoTrue
                .ForeColor.rgb = giRGBbluecolor3
                .Solid
            End With
            cht.SeriesCollection(3).Select
            With Selection.Format.Fill
                .Visible = msoTrue
                .ForeColor.rgb = giRGBbluecolor5
                .Solid
            End With
            cht.SeriesCollection(4).Select
            With Selection.Format.Fill
                .Visible = msoTrue
                .ForeColor.rgb = giRGBbluecolor7
                .Solid
            End With

        ElseIf seriescount = 5 Then
            cht.SeriesCollection(1).Select
            With Selection.Format.Fill
                .Visible = msoTrue
                .ForeColor.rgb = giRGBbluecolor1
                .Solid
            End With
            cht.SeriesCollection(2).Select
            With Selection.Format.Fill
                .Visible = msoTrue
                .ForeColor.rgb = giRGBbluecolor3
                .Solid
            End With
            cht.SeriesCollection(3).Select
            With Selection.Format.Fill
                .Visible = msoTrue
                .ForeColor.rgb = giRGBbluecolor5
                .Solid
            End With
            cht.SeriesCollection(4).Select
            With Selection.Format.Fill
                .Visible = msoTrue
                .ForeColor.rgb = giRGBbluecolor7
                .Solid
            End With
            cht.SeriesCollection(5).Select
            With Selection.Format.Fill
                .Visible = msoTrue
                .ForeColor.rgb = colorBlack
                .Solid
            End With

        ElseIf seriescount = 6 Then
            cht.SeriesCollection(1).Select
            With Selection.Format.Fill
                .Visible = msoTrue
                .ForeColor.rgb = giRGBbluecolor1
                .Solid
            End With
            cht.SeriesCollection(2).Select
            With Selection.Format.Fill
                .Visible = msoTrue
                .ForeColor.rgb = giRGBbluecolor2
                .Solid
            End With
            cht.SeriesCollection(3).Select
            With Selection.Format.Fill
                .Visible = msoTrue
                .ForeColor.rgb = giRGBbluecolor3
                .Solid
            End With
            cht.SeriesCollection(4).Select
            With Selection.Format.Fill
                .Visible = msoTrue
                .ForeColor.rgb = giRGBbluecolor4
                .Solid
            End With
            cht.SeriesCollection(5).Select
            With Selection.Format.Fill
                .Visible = msoTrue
                .ForeColor.rgb = giRGBbluecolor5
                .Solid
            End With
            cht.SeriesCollection(6).Select
            With Selection.Format.Fill
                .Visible = msoTrue
                .ForeColor.rgb = giRGBbluecolor6
                .Solid
            End With

        ElseIf seriescount > 6 Then
            If cht.HasTitle = True Then
                cht.ChartTitle.Text = "You have too many data series for this chart type. Please contact the Communications Department for further guidance."
            ElseIf cht.HasTitle = False Then
                'Add text box with text, size, and use correct font
                ''''Set txtB = cht.TextBoxes.Add(0, 0, 500, 40) '''' JP 2016 12 28 Shape not textbox
                Set txtB = cht.Shapes.AddTextbox(msoTextOrientationHorizontal, 0, 0, 500, 40)
                'Set txtB = cht.TextBoxes.Add(400, 100, 125, 20)
                '(horizontal placement, vertical placement, box width, box height)
                With txtB
                    .name = "TitleBox"
                    With .TextFrame2.TextRange
                        .Text = "You have too many data series for this chart type. Please contact the Communications Department for further guidance."
                        .Font.Size = 10
                        .Font.name = "Lato"
                        .Font.Fill.ForeColor.rgb = vbRed
                        'Left-align text in text box
                        .ParagraphFormat.Alignment = msoTextEffectAlignmentLeft
                    End With
                    .Fill.ForeColor.rgb = vbYellow
                End With
            End If

        End If

'    End If

End Sub

'JAS 2017
'JAS 2023
