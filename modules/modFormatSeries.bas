Attribute VB_Name = "modFormatSeries"
Option Explicit

' mode: "FILL", "LINE", or "BLUERAMP"
' fillTransparency: 0..1 (only used when mode="FILL")
' lineWeight: points (only used when mode="LINE")
Public Function FormatSeriesColors(cht As Chart, _
                                   ByVal mode As String, _
                                   Optional ByVal fillTransparency As Single = 0!, _
                                   Optional ByVal lineWeight As Single = 2!) As Boolean
    On Error GoTo CleanFail

    Dim n As Long, i As Long
    Dim palette(1 To 7) As Long
    Dim clr As Long

    If cht Is Nothing Then Exit Function

    ' --- Normalize mode
    mode = UCase$(Trim$(mode))
    If mode <> "FILL" And mode <> "LINE" And mode <> "BLUERAMP" Then
        MsgBox "Invalid mode. Use ""FILL"", ""LINE"", or ""BLUERAMP"".", vbExclamation, "FormatSeriesColors"
        Exit Function
    End If

    ' --- Blue ramp delegates to its own helper
    If mode = "BLUERAMP" Then
        ApplyBlueRampColors cht
        FormatSeriesColors = True
        Exit Function
    End If

    ' --- Build brand palette
    palette(1) = colorOcean
    palette(2) = colorCoral
    palette(3) = colorSky
    palette(4) = colorPine
    palette(5) = colorGold
    palette(6) = colorRust
    palette(7) = colorLavender

    n = cht.SeriesCollection.Count
    If n = 0 Then
        FormatSeriesColors = True
        Exit Function
    End If

    ' --- Apply colors
    For i = 1 To n
        clr = IIf(i <= 7, palette(i), colorSilver)  ' fallback for 8+

        With cht.SeriesCollection(i).Format
            If mode = "FILL" Then
                With .Fill
                    .Visible = msoTrue
                    .Solid
                    .ForeColor.rgb = clr
                    .transparency = IIf(fillTransparency < 0, 0, IIf(fillTransparency > 1, 1, fillTransparency))
                End With
            Else ' LINE
                With .Line
                    .Visible = msoTrue
                    .ForeColor.rgb = clr
                    .weight = lineWeight
                End With
            End If
        End With
    Next i

    FormatSeriesColors = True
    Exit Function

CleanFail:
    MsgError "FormatSeriesColors"
End Function


' Applies the ocean ramp palette with count-dependent color selection (1-6 series).
Private Sub ApplyBlueRampColors(cht As Chart)
    Dim n As Long, i As Long
    Dim ramp(1 To 6) As Long
    Dim txtB As Shape

    n = cht.SeriesCollection.Count
    If n = 0 Then Exit Sub

    If n > 6 Then
        Set txtB = cht.Shapes.AddTextbox(msoTextOrientationHorizontal, 0, 0, pieErrorBoxWidth, pieErrorBoxHeight)
        With txtB
            .Name = "TitleBox"
            With .TextFrame2.TextRange
                .Text = "You have too many data series for this chart type. Please contact the Communications Department for further guidance."
                .Font.Size = pieErrorFontSize
                .Font.Name = FontStyle
                .Font.Fill.ForeColor.rgb = vbRed
                .ParagraphFormat.Alignment = msoTextEffectAlignmentLeft
            End With
            .Fill.ForeColor.rgb = vbYellow
        End With
        Exit Sub
    End If

    Select Case n
        Case 1
            ramp(1) = rampOcean5
        Case 2
            ramp(1) = rampOcean5: ramp(2) = rampOcean2
        Case 3
            ramp(1) = rampOcean7: ramp(2) = rampOcean5: ramp(3) = rampOcean2
        Case 4
            ramp(1) = rampOcean7: ramp(2) = rampOcean5: ramp(3) = rampOcean3: ramp(4) = rampOcean1
        Case 5
            ramp(1) = colorBlack: ramp(2) = rampOcean7: ramp(3) = rampOcean5: ramp(4) = rampOcean3: ramp(5) = rampOcean1
        Case 6
            ramp(1) = rampOcean6: ramp(2) = rampOcean5: ramp(3) = rampOcean4: ramp(4) = rampOcean3: ramp(5) = rampOcean2: ramp(6) = rampOcean1
    End Select

    For i = 1 To n
        With cht.SeriesCollection(i).Format.Fill
            .Visible = msoTrue
            .Solid
            .ForeColor.rgb = ramp(i)
        End With
    Next i
End Sub
