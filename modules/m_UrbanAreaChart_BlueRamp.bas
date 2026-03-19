Attribute VB_Name = "m_UrbanAreaChart_BlueRamp"
Option Explicit



Sub UrbanAreaChart_BlueRamp()
'
    Dim seriescount As Long
    Dim cht As Chart
    Dim txtB As TextBox
    Dim bInsertLogo As Boolean
    Dim bInsertSource As Boolean
    Dim bFormatTitle As Boolean
    Dim bFormatGridlines As Boolean
    Dim bFormatXAxis As Boolean
    Dim bFormatXAxisTitle As Boolean
    Dim bFormatBlueFillColors As Boolean
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

        ActiveSheet.Shapes.AddChart2(-1, xlAreaStacked).Select
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
        bFormatBlueFillColors = FormatBlueFillColors(cht)

        'Position xaxis on tick marks
        ActiveChart.Axes(xlCategory).AxisBetweenCategories = False

        'Style xaxis tick marks
        ActiveChart.Axes(xlCategory).MajorTickMark = xlTickMarkOutside
        ActiveChart.Axes(xlCategory).MinorTickMark = xlTickMarkNone

'    End If

End Sub


Public Sub AreawithBlueRamp_onAction(control As IRibbonControl)
'
' Code for onAction callback. Ribbon control button
'
    Dim seriescount As Long
    Dim cht As Chart
    Dim txtB As TextBox
    Dim bInsertLogo As Boolean
    Dim bInsertSource As Boolean
    Dim bFormatTitle As Boolean
    Dim bFormatGridlines As Boolean
    Dim bFormatXAxis As Boolean
    Dim bFormatXAxisTitle As Boolean
    Dim bFormatBlueFillColors As Boolean
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

        ActiveSheet.Shapes.AddChart2(-1, xlAreaStacked).Select
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
        bFormatBlueFillColors = FormatBlueFillColors(cht)

        'Position xaxis on tick marks
        ActiveChart.Axes(xlCategory).AxisBetweenCategories = False

        'Style xaxis tick marks
        ActiveChart.Axes(xlCategory).MajorTickMark = xlTickMarkOutside
        ActiveChart.Axes(xlCategory).MinorTickMark = xlTickMarkNone

'    End If

End Sub

'JAS 2017
'JAS 2023
