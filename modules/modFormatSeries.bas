Attribute VB_Name = "modFormatSeries"
Option Explicit

' mode: "FILL" or "LINE"
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
    If mode <> "FILL" And mode <> "LINE" Then
        MsgBox "Invalid mode. Use ""FILL"" or ""LINE"".", vbExclamation, "FormatSeriesColors"
        Exit Function
    End If

    ' --- Build palette
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
        If i <= 7 Then
            clr = palette(i)
        Else
            clr = colorSilver   ' fallback for 8+
        End If

        With cht.SeriesCollection(i).Format
            If mode = "FILL" Then
                With .Fill
                    .Visible = msoTrue
                    .ForeColor.rgb = clr
                    .transparency = IIf(fillTransparency < 0, 0, IIf(fillTransparency > 1, 1, fillTransparency))
                    .Solid
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

