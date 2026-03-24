Attribute VB_Name = "modChartLine"
Option Explicit

Private Sub BuildLineChart()
    On Error GoTo CleanFail

    Dim cht As Chart

    Set cht = GetTargetChart(xlLine)
    If cht Is Nothing Then Exit Sub

    ' Shared formatting pipeline
    ApplyChartPipeline cht, "LINE"
    Call RemoveShadow(cht)

    ' Line-specific: axis starts on first data point (not between categories)
    If cht.HasAxis(xlCategory) Then
        cht.Axes(xlCategory).AxisBetweenCategories = False
        cht.Axes(xlCategory).MajorTickMark = xlTickMarkOutside
        cht.Axes(xlCategory).MinorTickMark = xlTickMarkNone
        ' Re-hide axis line: AxisBetweenCategories assignment can re-show it
        cht.Axes(xlCategory).Select
        Selection.Format.Line.Visible = msoFalse
    End If

    If cht.HasAxis(xlValue) Then
        cht.Axes(xlValue).MajorTickMark = xlTickMarkOutside
        cht.Axes(xlValue).MinorTickMark = xlTickMarkNone
        cht.Axes(xlValue).Format.Line.Visible = msoFalse
    End If
    Exit Sub
CleanFail:
    MsgError "BuildLineChart"
End Sub

Sub LineChart()
    BuildLineChart
End Sub
