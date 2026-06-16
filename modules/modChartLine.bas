Attribute VB_Name = "modChartLine"
Option Explicit

Private Sub BuildLineChart()
    On Error GoTo CleanFail
    AppFast

    Dim cht As Chart

    Set cht = GetTargetChart(xlLine)
    If cht Is Nothing Then GoTo CleanExit

    ' Shared formatting pipeline with line-specific defaults
    ApplyChartPipeline cht, "LINE", LineChartDefaults()
    Call RemoveShadow(cht)

    ' Line-specific: axis starts on first data point (not between categories)
    If cht.HasAxis(xlCategory) Then
        cht.Axes(xlCategory).AxisBetweenCategories = False
        cht.Axes(xlCategory).MajorTickMark = xlTickMarkOutside
        cht.Axes(xlCategory).MinorTickMark = xlTickMarkNone
        ' Re-format axis line to white: AxisBetweenCategories assignment can re-show it
        FormatAxisLineWhite cht.Axes(xlCategory)
    End If

    If cht.HasAxis(xlValue) Then
        cht.Axes(xlValue).MajorTickMark = xlTickMarkOutside
        cht.Axes(xlValue).MinorTickMark = xlTickMarkNone
        FormatAxisLineWhite cht.Axes(xlValue)
    End If
CleanExit:
    AppRestore
    Exit Sub
CleanFail:
    AppRestore
    MsgError "BuildLineChart"
End Sub

Sub LineChart()
    BuildLineChart
End Sub
