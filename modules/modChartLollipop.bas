Attribute VB_Name = "modChartLollipop"
'==== Module: modChartLollipop ====
' Builds a lollipop chart by wrapping the bar chart pipeline and converting each
' series to a stick-and-dot via horizontal error bars.
'
' Technique: horizontal error bars styled as stems with an oval arrowhead ("candy") at the data-value end.
'   1. Build a standard horizontal bar chart (full pipeline via BarChart)
'   2. For each series, add a horizontal error bar: Minus direction, No Cap, 100%
'      This extends a line from the data value back to zero.
'   3. Set bar fill to No Fill — the bar becomes invisible; only the error bar line shows.
'   4. Format the error bar line with an oval arrowhead at the value end (the "candy").
'      The line itself becomes the stick.
Option Explicit

Private Sub BuildLollipopChart()
    Dim cht As Chart
    Dim srs As Series
    Dim n As Long, i As Long
    Dim clr As Long

    ' Create and pipeline-format a bar chart, then convert to lollipop style
    BarChart

    Set cht = ActiveChart
    If cht Is Nothing Then Exit Sub

    n = cht.SeriesCollection.Count

    cht.ChartGroups(1).GapWidth = lollipopGapWidth

    For i = 1 To n
        Set srs = cht.SeriesCollection(i)
        ' Delegate to modFormatSeries so palette toggle is respected
        clr = GetPaletteColor(i)

        ' Hide the bar — fill and border both invisible
        With srs.Format.Fill
            .Visible = msoFalse
        End With
        With srs.Format.Line
            .Visible = msoFalse
        End With

        ' Add horizontal error bar extending from the value back to zero (the stick)
        srs.ErrorBar Direction:=xlX, Include:=xlMinusValues, Type:=xlErrorBarTypePercent, Amount:=100
        srs.ErrorBars.EndStyle = xlNoCap

        ' Format stick: brand colour, round join, oval arrowhead at the value end (the candy)
        srs.ErrorBars.Select
        With Selection.Format.Line
            .Visible = msoTrue
            .ForeColor.RGB = clr
            .Weight = lollipopStickWeight
            .JoinType = msoLineJoinRound
            .BeginArrowheadStyle = msoArrowheadOval
            .BeginArrowheadLength = msoArrowheadLengthMedium
            .BeginArrowheadWidth = msoArrowheadWidthMedium
        End With
    Next i
End Sub


Sub LollipopChart()
    BuildLollipopChart
End Sub
