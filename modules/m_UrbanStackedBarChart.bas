Attribute VB_Name = "m_UrbanStackedBarChart"
Option Explicit



Sub UrbanStackedBarChart()
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
    Dim bSetWebVersion As Boolean

'    'Check ActiveChart
'    If ActiveChart Is Nothing Then
'        MsgBox "First, create a chart, or select an existing chart, and try again.", vbExclamation, "No Active Chart"
'    Else

        '''bSetWebVersion = SetWebVersion(gWebCancel)
        SetWebVersion_NEW
        If gWebCancel Then
            Exit Sub
        End If

        ActiveSheet.Shapes.AddChart2(-1, xlBarStacked).Select
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

        'Set series overlap to 0% and gap width to 70%
        ActiveChart.ChartGroups(1).Overlap = 100
        ActiveChart.ChartGroups(1).GapWidth = 70

'    End If

End Sub

Public Sub StackedBar_onAction(control As IRibbonControl)
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
    Dim bSetWebVersion As Boolean

'    'Check ActiveChart
'    If ActiveChart Is Nothing Then
'        MsgBox "First, create a chart, or select an existing chart, and try again.", vbExclamation, "No Active Chart"
'    Else

        ''''bSetWebVersion = SetWebVersion(gWebCancel)
        SetWebVersion_NEW
        If gWebCancel Then
            Exit Sub
        End If

        ActiveSheet.Shapes.AddChart2(-1, xlBarStacked).Select
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

        'Set series overlap to 0% and gap width to 70%
        ActiveChart.ChartGroups(1).Overlap = 100
        ActiveChart.ChartGroups(1).GapWidth = 70

'    End If

End Sub

'JAS 2017
'JAS 2023
