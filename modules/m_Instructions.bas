Attribute VB_Name = "m_Instructions"
Option Explicit

Sub NotesButton()

    Dim answer As VbMsgBoxResult
    Dim mystring As String

   mystring = _
        "This add-in was created for INSO staff to create and style " _
             & "their graphs according to the INSO data visualisation standards." & vbCrLf & _
               "-This add-in creates a chart for a selected table or data range and formats; " _
             & "according to the INSO standards. you still need to write your own title," & vbCrLf & _
               "subtitle, axis titles (of applicable) and source/note text (if applicable). " _
             & "-If you have questions about this add-in or how to create or style your graphs, " _
             & "please contact the HQ Senior Information Team." & vbCrLf & _
               "-This is v0.9 of this add-in."

    answer = MsgBox(mystring, vbInformation, "Notes")

End Sub

Public Sub NotesButton_onAction(control As IRibbonControl)
'
' Code for onAction callback. Ribbon control button
'
    Dim answer As VbMsgBoxResult
    Dim mystring As String
    
    mystring = _
        "This add-in was created for INSO staff to create and style " _
             & "their graphs according to the INSO data visualisation standards." & vbCrLf & _
               "-This add-in creates a chart for a selected table or data range and formats; " _
             & "according to the INSO standards. you still need to write your own title," & vbCrLf & _
               "subtitle, axis titles (of applicable) and source/note text (if applicable). " _
             & "-If you have questions about this add-in or how to create or style your graphs, " _
             & "please contact the HQ Senior Information Team." & vbCrLf & _
               "-This is v0.9 of this add-in."

    answer = MsgBox(mystring, vbInformation, "Notes")

End Sub

