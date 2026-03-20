Attribute VB_Name = "modFormatLineFill"
'==== Module: modFormatLineFill ====
Option Explicit

'   FILL
Public Sub ApplyFill(ByVal colorRGB As Long, Optional ByVal transparency As Single = 0!)
    ' transparency: 0 = opaque, 1 = fully transparent (chart model)
    On Error GoTo CleanFail

    Dim tgt As Object
    Set tgt = GetFillTarget()
    If tgt Is Nothing Then
        MsgSelectTarget
        Exit Sub
    End If

    With tgt.Format.Fill
        .Visible = msoTrue
        .Solid
        .ForeColor.rgb = colorRGB

        ' Charts: .Transparency 0-1
        ' Shapes:   .Transparency 0-1 (Excel normalizes)
        If transparency < 0 Then transparency = 0
        If transparency > 1 Then transparency = 1
        .transparency = transparency
    End With

    Exit Sub

CleanFail:
    MsgError "ApplyFill"
End Sub


Public Sub RemoveFill()
    On Error GoTo CleanFail

    Dim tgt As Object
    Set tgt = GetFillTarget()
    If tgt Is Nothing Then
        MsgSelectTarget
        Exit Sub
    End If

    tgt.Format.Fill.Visible = msoFalse
    Exit Sub

CleanFail:
    MsgError "RemoveFill"
End Sub

'   TARGET DETECTION HELPERS
Private Function GetFillTarget() As Object
    If Not ActiveChart Is Nothing Then
        If Not Selection Is Nothing Then
            If HasFillFormat(Selection) Then
                Set GetFillTarget = Selection
                Exit Function
            End If
        End If
        Set GetFillTarget = ActiveChart.ChartArea
        Exit Function
    End If

    If Not Selection Is Nothing Then
        If HasFillFormat(Selection) Then Set GetFillTarget = Selection
    End If
End Function


Private Function HasFillFormat(o As Object) As Boolean
    On Error Resume Next
    Dim x: Set x = o.Format.Fill
    HasFillFormat = (Err.Number = 0)
    Err.Clear
End Function
