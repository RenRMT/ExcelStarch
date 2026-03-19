Attribute VB_Name = "modRemoveLegendResize"
Option Explicit

Private Sub BuildRemoveLegendResize()
    If ActiveChart Is Nothing Then
        MsgNoActiveChart
        Exit Sub
    End If

    RemoveLegendandResize ActiveChart
End Sub

Private Sub RemoveLegendandResize(cht As Chart)
    If cht.HasLegend Then cht.Legend.Delete

    cht.PlotArea.Select
    Selection.Height = removeLegend_webHeight
    Selection.Top = removeLegend_webTop
    Selection.Width = removeLegend_webWidth
    Selection.Left = removeLegend_webLeft
End Sub


Sub RemoveLegendResizeButton()
    BuildRemoveLegendResize
End Sub

Public Sub RemoveLegendResizeButton_onAction(control As IRibbonControl)
    BuildRemoveLegendResize
End Sub
