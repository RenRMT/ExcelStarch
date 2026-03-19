Attribute VB_Name = "m_UrbanRemoveLegendandResize"
Option Explicit



Sub UrbanRemoveLegendandResize()

    Dim cht As Chart
    Dim bRemoveLegendandResize As Boolean

    'Check ActiveChart
    If ActiveChart Is Nothing Then
        MsgBox "First, create a chart, or select an existing chart, and try again.", vbExclamation, "No Active Chart"
    Else

        If SetWebVersionRemoveLegend = vbCancel Then Exit Sub

        'SetWebVersionRemoveLegend

        bRemoveLegendandResize = RemoveLegendandResize(cht)

    End If

End Sub

Public Sub RemoveLegendResizeButton_onAction(control As IRibbonControl)
'
' Code for onAction callback. Ribbon control button
'
    Dim cht As Chart
    Dim bRemoveLegendandResize As Boolean

    'Check ActiveChart
    If ActiveChart Is Nothing Then
        MsgBox "First, create a chart, or select an existing chart, and try again.", vbExclamation, "No Active Chart"
    Else

        If SetWebVersionRemoveLegend = vbCancel Then Exit Sub

        'SetWebVersionRemoveLegend

        bRemoveLegendandResize = RemoveLegendandResize(cht)

    End If

End Sub

'JAS 2017
'JAS 2023
