Attribute VB_Name = "modMessages"
'==== Module: modMessages ====
' Centralised user-facing messages. Call these instead of inline MsgBox.
Option Explicit

' Guard: no chart selected
Public Sub MsgNoActiveChart()
    MsgBox "Select a chart and try again.", vbExclamation, "No Active Chart"
End Sub

' Guard: no valid chart element or shape selected
Public Sub MsgSelectTarget()
    MsgBox "Select a chart element or shape.", vbExclamation, "No Selection"
End Sub

' Generic error handler — call from a CleanFail label while Err object is populated
Public Sub MsgError(ByVal source As String)
    MsgBox source & ": " & Err.Number & " - " & Err.Description, vbExclamation
End Sub

' Adds a styled error box to the chart when the series count exceeds the supported limit.
' Overlays a yellow/red warning directly on the chart area.
Public Sub MsgTooManySeries(cht As Chart)
    Dim txtB As Shape
    Set txtB = cht.Shapes.AddTextbox(msoTextOrientationHorizontal, 0, 0, errorBoxWidth, errorBoxHeight)
    With txtB
        .Name = "ErrorBox"
        With .TextFrame2.TextRange
            .Text = "You have too many data series for this chart type. Please contact the Communications Department for further guidance."
            .Font.Size = errorBoxFontSize
            .Font.Name = gsPRIMARY_FONT
            .Font.Fill.ForeColor.rgb = vbRed
            .ParagraphFormat.Alignment = msoTextEffectAlignmentLeft
        End With
        .Fill.ForeColor.rgb = vbYellow
    End With
End Sub
