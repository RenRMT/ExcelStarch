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
    MsgBox "Select a chart element or shape.", vbExclamation
End Sub

' Generic error handler — call from a CleanFail label while Err object is populated
Public Sub MsgError(ByVal source As String)
    MsgBox source & ": " & Err.Number & " - " & Err.Description, vbExclamation
End Sub
