Attribute VB_Name = "modInstructions"
Option Explicit

Private Sub ShowNotes()
    Dim mystring As String

    mystring = _
        "This add-in was created for " & orgName & " staff to create and style " _
             & "their graphs according to the " & orgName & " data visualisation standards." & vbCrLf & _
               "-This add-in creates a chart for a selected table or data range and formats; " _
             & "according to the " & orgName & " standards. you still need to write your own title," & vbCrLf & _
               "subtitle, axis titles (of applicable) and source/note text (if applicable). " _
             & "-If you have questions about this add-in or how to create or style your graphs, " _
             & "please contact the " & orgSupportContact & "." & vbCrLf & _
               "-This is " & exportAddInVersion & " of this add-in."

    MsgBox mystring, vbInformation, "Notes"
End Sub

Sub NotesButton()
    ShowNotes
End Sub
