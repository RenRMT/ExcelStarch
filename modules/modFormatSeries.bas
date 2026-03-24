Attribute VB_Name = "modFormatSeries"
Option Explicit

' Palette order toggle. False = default, True = alternative.
' Default:     colorData1, colorData2, colorData3, colorData4, colorData5, colorData6, colorData7
' Alternative: colorData1, colorData7, colorData3, colorData4, colorData5, colorData2, colorData6
Private m_useAltOrder As Boolean

Public Function GetPaletteColor(ByVal i As Long) As Long
    ' Returns the brand colour for series index i, respecting the current palette order.
    ' Falls back to colorNeutral1 for i > 7.
    Dim palette(1 To 7) As Long
    If m_useAltOrder Then
        palette(1) = colorData1:    palette(2) = colorData7:    palette(3) = colorData3
        palette(4) = colorData4:    palette(5) = colorData5:    palette(6) = colorData2
        palette(7) = colorData6
    Else
        palette(1) = colorData1:    palette(2) = colorData2:    palette(3) = colorData3
        palette(4) = colorData4:    palette(5) = colorData5:    palette(6) = colorData6
        palette(7) = colorData7
    End If
    GetPaletteColor = IIf(i >= 1 And i <= 7, palette(i), colorNeutral1)
End Function

Public Sub TogglePaletteOrder()
    m_useAltOrder = Not m_useAltOrder
    MsgPaletteOrderToggled m_useAltOrder
End Sub

' mode: "FILL" or "LINE"
' fillTransparency: 0..1 (only used when mode="FILL")
' lineWeight: points (only used when mode="LINE")
Public Function FormatSeriesColors(cht As Chart, _
                                   ByVal mode As String, _
                                   Optional ByVal fillTransparency As Single = 0!, _
                                   Optional ByVal lineWeight As Single = 2!) As Boolean
    On Error GoTo CleanFail

    Dim n As Long, i As Long
    Dim clr As Long

    If cht Is Nothing Then Exit Function

    ' --- Normalize mode
    mode = UCase$(Trim$(mode))
    If mode <> "FILL" And mode <> "LINE" Then
        MsgInvalidColorMode
        Exit Function
    End If

    n = cht.SeriesCollection.Count
    If n = 0 Then
        FormatSeriesColors = True
        Exit Function
    End If

    ' --- Apply colors
    For i = 1 To n
        clr = GetPaletteColor(i)

        With cht.SeriesCollection(i).Format
            If mode = "FILL" Then
                With .Fill
                    .Visible = msoTrue
                    .Solid
                    .ForeColor.RGB = clr
                    .transparency = IIf(fillTransparency < 0, 0, IIf(fillTransparency > 1, 1, fillTransparency))
                End With
            Else ' LINE
                With .Line
                    .Visible = msoTrue
                    .ForeColor.RGB = clr
                    .Weight = lineWeight
                End With
            End If
        End With
    Next i

    FormatSeriesColors = True
    Exit Function

CleanFail:
    MsgError "FormatSeriesColors"
End Function
