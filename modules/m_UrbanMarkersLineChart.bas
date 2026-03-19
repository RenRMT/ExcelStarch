Attribute VB_Name = "m_UrbanMarkersLineChart"
Option Explicit



Sub UrbanMarkersLineChart()
'
    Dim seriescount As Long
    Dim imarker As Long
    Dim cht As Chart
    Dim txtB As TextBox
    Dim bInsertLogo As Boolean
    Dim bInsertSource As Boolean
    Dim bFormatTitle As Boolean
    Dim bFormatGridlines As Boolean
    Dim bFormatXAxis As Boolean
    Dim bFormatXAxisTitle As Boolean
    Dim bFormatLineColors As Boolean
    Dim bOuterFormat As Boolean

'    'Check ActiveChart
'    If ActiveChart Is Nothing Then
'        MsgBox "First, create a chart, or select an existing chart, and try again.", vbExclamation, "No Active Chart"
'    Else

        ''bSetWebVersion = SetWebVersion(gWebCancel)
        SetWebVersion_NEW
        If gWebCancel Then
            Exit Sub
        End If

        ActiveSheet.Shapes.AddChart2(-1, xlLineMarkers).Select
        'Duplicate the selected chart
        'ActiveChart.Parent.Duplicate.Select

        Set cht = ActiveChart

        bOuterFormat = OuterFormat(cht)
        bFormatXAxisTitle = FormatXAxisTitle(cht)
        bInsertLogo = InsertLogo(cht)
        bInsertSource = InsertSource(cht)
        bFormatTitle = FormatTitle(cht)
        bFormatGridlines = FormatGridlines(cht)
        bFormatXAxis = FormatXAxis(cht)
        bFormatLineColors = FormatSeriesColors(cht, "LINE")

        'Position xaxis on tick marks
        cht.Axes(xlCategory).AxisBetweenCategories = False

        'Style xaxis and yaxis tick marks
        cht.Axes(xlCategory).MajorTickMark = xlTickMarkOutside
        cht.Axes(xlCategory).MinorTickMark = xlTickMarkNone
        cht.Axes(xlValue).MajorTickMark = xlTickMarkOutside
        cht.Axes(xlValue).MinorTickMark = xlTickMarkNone

        'Change colors of lines
        With cht
            seriescount = .SeriesCollection.Count
        End With

        'Change marker style and size
        For imarker = 1 To seriescount
            cht.SeriesCollection(imarker).Select
            With Selection
                .MarkerStyle = xlMarkerStyleCircle
                .MarkerSize = 7
            End With
            With Selection.Format.Fill
                .Visible = msoTrue
                .ForeColor.rgb = giRGBwhitecolor
            End With
        Next

    'End If

End Sub

Public Sub LinewithMarkers_onAction(control As IRibbonControl)
'
' Code for onAction callback. Ribbon control button
'
    Dim seriescount As Long
    Dim imarker As Long
    Dim cht As Chart
    Dim txtB As TextBox
    Dim bInsertLogo As Boolean
    Dim bInsertSource As Boolean
    Dim bFormatTitle As Boolean
    Dim bFormatGridlines As Boolean
    Dim bFormatXAxis As Boolean
    Dim bFormatXAxisTitle As Boolean
    Dim bFormatLineColors As Boolean
    Dim bOuterFormat As Boolean
    Dim bSetWebVersion As Boolean

    'Check ActiveChart
'    If ActiveChart Is Nothing Then
'        MsgBox "First, create a chart, or select an existing chart, and try again.", vbExclamation, "No Active Chart"
'    Else

        ''''bSetWebVersion = SetWebVersion(gWebCancel)
        SetWebVersion_NEW
        If gWebCancel Then
            Exit Sub
        End If

        ActiveSheet.Shapes.AddChart2(-1, xlLineMarkers).Select
        'Duplicate the selected chart
        'ActiveChart.Parent.Duplicate.Select
        
        Set cht = ActiveChart

        bOuterFormat = OuterFormat(cht)
        bFormatXAxisTitle = FormatXAxisTitle(cht)
        bInsertLogo = InsertLogo(cht)
        bInsertSource = InsertSource(cht)
        bFormatTitle = FormatTitle(cht)
        bFormatGridlines = FormatGridlines(cht)
        bFormatXAxis = FormatXAxis(cht)
        bFormatLineColors = FormatSeriesColors(cht, "LINE")

        'Position xaxis on tick marks
        cht.Axes(xlCategory).AxisBetweenCategories = False

        'Style xaxis and yaxis tick marks
        cht.Axes(xlCategory).MajorTickMark = xlTickMarkOutside
        cht.Axes(xlCategory).MinorTickMark = xlTickMarkNone
        cht.Axes(xlValue).MajorTickMark = xlTickMarkOutside
        cht.Axes(xlValue).MinorTickMark = xlTickMarkNone

        'Change colors of lines
        With cht
            seriescount = .SeriesCollection.Count
        End With

        'Change marker style and size
        For imarker = 1 To seriescount
            cht.SeriesCollection(imarker).Select
            With Selection
                .MarkerStyle = xlMarkerStyleCircle
                .MarkerSize = 7
            End With
            With Selection.Format.Fill
                .Visible = msoTrue
                .ForeColor.rgb = giRGBwhitecolor
            End With
        Next

    'End If

End Sub

'JAS 2017
'JAS 2023
