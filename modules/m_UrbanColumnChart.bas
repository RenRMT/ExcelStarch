Attribute VB_Name = "m_UrbanColumnChart"
Option Explicit

Public Sub Column_onAction(control As IRibbonControl)
'
' Code for onAction callback. Ribbon control button
'
    Dim seriescount As Long
    Dim ishadow As Long
    Dim cht As Chart
    Dim txtB As TextBox
    Dim bInsertLogo As Boolean
    Dim bInsertSource As Boolean
    Dim bFormatTitle As Boolean
    Dim bFormatGridlines As Boolean
    Dim bFormatXAxis As Boolean
    Dim bFormatXAxisTitle As Boolean
    Dim bFormatFillColors As Boolean
    Dim bOuterFormat As Boolean
    Dim bRemoveShadow As Boolean
    Dim bSetWebVersion As Boolean

'    'Check ActiveChart
'    If ActiveChart Is Nothing Then
'        MsgBox "First, create a chart, or select an existing chart, and try again.", vbExclamation, "No Active Chart"
'    Else

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
        bFormatFillColors = FormatSeriesColors(cht, "FILL")

        'Style xaxis tick marks
        ActiveChart.Axes(xlCategory).MajorTickMark = xlTickMarkNone
        ActiveChart.Axes(xlCategory).MinorTickMark = xlTickMarkNone

        'Set series overlap and gap width. Change default setting on modConfig
        ActiveChart.ChartGroups(1).Overlap = seriesOverlap
        ActiveChart.ChartGroups(1).GapWidth = seriesGapWidth

'    End If

End Sub
