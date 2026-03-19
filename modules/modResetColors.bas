Attribute VB_Name = "modResetColors"
Option Explicit

' Core: make all series on the chart gray (line + fill).
' - If cht is Nothing, falls back to ActiveChart.
' - If duplicateChart := True, duplicates the chart first and then grays out the duplicate.
' - grayColor defaults to colorSilver if omitted.
Public Sub GrayOutChart(Optional ByVal cht As Chart = Nothing, _
                        Optional ByVal duplicateChart As Boolean = True, _
                        Optional ByVal grayColor As Long = 0)
    On Error GoTo CleanFail

    Dim targetChart As Chart
    Dim i As Long, n As Long

    ' Resolve chart
    If cht Is Nothing Then
        If ActiveChart Is Nothing Then
            MsgNoActiveChart
            Exit Sub
        End If
        Set targetChart = ActiveChart
    Else
        Set targetChart = cht
    End If

    ' Default gray color
    If grayColor = 0 Then grayColor = colorSilver  ' or giRGBgridlinesprint if you prefer

    ' Confirm once
    Dim answer As VbMsgBoxResult
    answer = MsgBox(IIf(duplicateChart, _
                       "This will duplicate your chart and make the duplicate gray.", _
                       "This will make your current chart gray."), _
                       vbExclamation + vbOKCancel)
    If answer <> vbOK Then Exit Sub

    ' Duplicate if requested
    If duplicateChart Then
        targetChart.Parent.Duplicate.Select
        Set targetChart = ActiveChart
        If targetChart Is Nothing Then
            MsgBox "Could not resolve duplicated chart.", vbExclamation
            Exit Sub
        End If
    End If
    ' Gray out each series (line + fill), without Selection
    n = targetChart.SeriesCollection.Count
    For i = 1 To n
        With targetChart.SeriesCollection(i).Format
            With .Line
                .Visible = msoTrue
                .ForeColor.rgb = grayColor
            End With
            With .Fill
                .Visible = msoTrue
                .ForeColor.rgb = grayColor
                .Solid
            End With
        End With
    Next i

    On Error GoTo 0

    Exit Sub

CleanFail:
    MsgError "GrayOutChart"
End Sub


' macro entry point: duplicates & grays using colorSilver
Public Sub StartWithGray()
    GrayOutChart cht:=Nothing, duplicateChart:=True, grayColor:=colorSilver
End Sub





