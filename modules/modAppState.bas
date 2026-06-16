Attribute VB_Name = "modAppState"
'==== Module: modAppState ====
' Centralised Excel application-state management for the chart operations.
'
' Purpose
' -------
' Chart builders and tools perform many Excel object-model writes per action
' (shapes, textboxes, series colours, logos). Without suppressing screen
' updating this causes visible flicker and slower runs. CLAUDE.md Core
' Principle #6 ("Always restore application state") requires that whatever we
' disable is reliably re-enabled — including on error paths.
'
' Re-entrancy
' -----------
' Some operations nest (e.g. BuildLollipopChart calls BarChart). A plain
' on/off flag would let an inner AppRestore re-enable screen updating while
' the outer operation is still running. A depth counter solves this: screen
' updating is captured/disabled only on the outermost AppFast, and restored
' only when the matching outermost AppRestore is reached.
'
' Usage
' -----
'   Public Sub BuildXxx()
'       On Error GoTo CleanFail
'       AppFast
'       ' ... heavy object-model work ...
'       AppRestore            ' normal exit
'       Exit Sub
'   CleanFail:
'       AppRestore            ' guaranteed restore on error
'       MsgError "BuildXxx"
'   End Sub
'
' Always pair every AppFast with an AppRestore on BOTH the normal exit and the
' error handler. Calls do not have to balance perfectly — AppReset is available
' as a hard reset if a depth leak is ever suspected.
Option Explicit

Private mDepth As Long
Private mSavedScreenUpdating As Boolean

' Disable screen updating for the duration of an operation. Safe to nest.
Public Sub AppFast()
    If mDepth = 0 Then
        mSavedScreenUpdating = Application.ScreenUpdating
        Application.ScreenUpdating = False
    End If
    mDepth = mDepth + 1
End Sub

' Restore screen updating once the outermost operation completes.
Public Sub AppRestore()
    If mDepth > 0 Then mDepth = mDepth - 1
    If mDepth = 0 Then
        Application.ScreenUpdating = mSavedScreenUpdating
    End If
End Sub

' Hard reset — unconditionally re-enable screen updating and clear the depth
' counter. Use only as a recovery measure (e.g. from the Immediate Window) if
' an unbalanced AppFast/AppRestore is ever suspected.
Public Sub AppReset()
    mDepth = 0
    Application.ScreenUpdating = True
End Sub
