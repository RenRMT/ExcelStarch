Attribute VB_Name = "modChartMarkersLine"
Option Explicit

Private Sub BuildMarkersLineChart()
    Dim seriescount As Long
    Dim imarker As Long
    Dim cht As Chart

    ActiveSheet.Shapes.AddChart2(-1, xlLineMarkers).Select
    ActiveChart.Parent.Duplicate.Select

    Set cht = ActiveChart
    If cht Is Nothing Then Exit Sub

    ' Full pipeline
    ApplyChartPipeline cht, "LINE"

    cht.Axes(xlCategory).AxisBetweenCategories = False
    cht.Axes(xlCategory).MajorTickMark = xlTickMarkOutside
    cht.Axes(xlCategory).MinorTickMark = xlTickMarkNone
    cht.Axes(xlValue).MajorTickMark = xlTickMarkOutside
    cht.Axes(xlValue).MinorTickMark = xlTickMarkNone

    ' Markers: circle, white fill
    seriescount = cht.SeriesCollection.Count
    For imarker = 1 To seriescount
        With cht.SeriesCollection(imarker)
            .MarkerStyle = xlMarkerStyleCircle
            .MarkerSize = scatterMarkerSize
            With .Format.Fill
                .Visible = msoTrue
                .Solid
                .ForeColor.rgb = colorWhite
            End With
        End With
    Next imarker
End Sub

Sub MarkersLineChart()
    BuildMarkersLineChart
End Sub
