Attribute VB_Name = "m_CommonBlueFillColorFunctions"
Option Explicit

Function FormatBlueFillColors(cht As Chart) As Boolean
    Dim seriescount As Long
    Dim txtB As Shape    '''' (not TextBox)

    Set cht = ActiveChart

    'Change colors of areas
    With cht
        seriescount = .SeriesCollection.Count
    End With

    If seriescount = 1 Then
        cht.SeriesCollection(1).Select
        With Selection.Format.Fill
            .Visible = msoTrue
            .ForeColor.rgb = rampOcean5
            .Solid
        End With

    ElseIf seriescount = 2 Then
        cht.SeriesCollection(1).Select
        With Selection.Format.Fill
            .Visible = msoTrue
            .ForeColor.rgb = rampOcean5
            .Solid
        End With
        cht.SeriesCollection(2).Select
        With Selection.Format.Fill
            .Visible = msoTrue
            .ForeColor.rgb = rampOcean2
            .Solid
        End With

    ElseIf seriescount = 3 Then
        cht.SeriesCollection(1).Select
        With Selection.Format.Fill
            .Visible = msoTrue
            .ForeColor.rgb = rampOcean7
            .Solid
        End With
        cht.SeriesCollection(2).Select
        With Selection.Format.Fill
            .Visible = msoTrue
            .ForeColor.rgb = rampOcean5
            .Solid
        End With
        cht.SeriesCollection(3).Select
        With Selection.Format.Fill
            .Visible = msoTrue
            .ForeColor.rgb = rampOcean2
            .Solid
        End With

    ElseIf seriescount = 4 Then
        cht.SeriesCollection(1).Select
        With Selection.Format.Fill
            .Visible = msoTrue
            .ForeColor.rgb = rampOcean7
            .Solid
        End With
        cht.SeriesCollection(2).Select
        With Selection.Format.Fill
            .Visible = msoTrue
            .ForeColor.rgb = rampOcean5
            .Solid
        End With
        cht.SeriesCollection(3).Select
        With Selection.Format.Fill
            .Visible = msoTrue
            .ForeColor.rgb = rampOcean3
            .Solid
        End With
        cht.SeriesCollection(4).Select
        With Selection.Format.Fill
            .Visible = msoTrue
            .ForeColor.rgb = rampOcean1
            .Solid
        End With

    ElseIf seriescount = 5 Then
        cht.SeriesCollection(1).Select
        With Selection.Format.Fill
            .Visible = msoTrue
            .ForeColor.rgb = colorBlack
            .Solid
        End With
        cht.SeriesCollection(2).Select
        With Selection.Format.Fill
            .Visible = msoTrue
            .ForeColor.rgb = rampOcean7
            .Solid
        End With
        cht.SeriesCollection(3).Select
        With Selection.Format.Fill
            .Visible = msoTrue
            .ForeColor.rgb = rampOcean5
            .Solid
        End With
        cht.SeriesCollection(4).Select
        With Selection.Format.Fill
            .Visible = msoTrue
            .ForeColor.rgb = rampOcean3
            .Solid
        End With
        cht.SeriesCollection(5).Select
        With Selection.Format.Fill
            .Visible = msoTrue
            .ForeColor.rgb = rampOcean1
            .Solid
        End With

    ElseIf seriescount = 6 Then
        cht.SeriesCollection(1).Select
        With Selection.Format.Fill
            .Visible = msoTrue
            .ForeColor.rgb = rampOcean6
            .Solid
        End With
        cht.SeriesCollection(2).Select
        With Selection.Format.Fill
            .Visible = msoTrue
            .ForeColor.rgb = rampOcean5
            .Solid
        End With
        cht.SeriesCollection(3).Select
        With Selection.Format.Fill
            .Visible = msoTrue
            .ForeColor.rgb = rampOcean4
            .Solid
        End With
        cht.SeriesCollection(4).Select
        With Selection.Format.Fill
            .Visible = msoTrue
            .ForeColor.rgb = rampOcean3
            .Solid
        End With
        cht.SeriesCollection(5).Select
        With Selection.Format.Fill
            .Visible = msoTrue
            .ForeColor.rgb = rampOcean2
            .Solid
        End With
        cht.SeriesCollection(6).Select
        With Selection.Format.Fill
            .Visible = msoTrue
            .ForeColor.rgb = rampOcean1
            .Solid
        End With


    ElseIf seriescount > 6 Then
        If cht.HasTitle = True Then
            cht.ChartTitle.Text = "You have too many data series for this chart type."
        ElseIf cht.HasTitle = False Then
            'Add text box with text, size, and use correct font
            ''''Set txtB = cht.TextBoxes.Add(0, 0, 500, 40) '''' JP 2016 12 28 Shape not textbox
            Set txtB = cht.Shapes.AddTextbox(msoTextOrientationHorizontal, 0, 0, 500, 40)
            'Set txtB = cht.TextBoxes.Add(400, 100, 125, 20)
            '(horizontal placement, vertical placement, box width, box height)
            With txtB
                .name = "TitleBox"
                With .TextFrame2.TextRange
                    .Text = "You have too many data series for this chart type."
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

End Function

