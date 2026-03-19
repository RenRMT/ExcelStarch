Attribute VB_Name = "modFormatLineFill"
'==== Module: modFormatLineFill ====
Option Explicit

'   OUTLINE (LINE)
Public Sub ApplyOutline(ByVal colorRGB As Long, Optional ByVal weightPts As Single = 2!)
    On Error GoTo CleanFail

    Dim tgt As Object
    Set tgt = GetLineTarget()
    If tgt Is Nothing Then
        MsgSelectTarget
        Exit Sub
    End If

    With tgt.Format.Line
        .Visible = msoTrue
        .ForeColor.rgb = colorRGB
        .weight = weightPts
    End With
    Exit Sub

CleanFail:
    MsgError "ApplyOutline"
End Sub


Public Sub RemoveOutline()
    On Error GoTo CleanFail

    Dim tgt As Object
    Set tgt = GetLineTarget()
    If tgt Is Nothing Then
        MsgSelectTarget
        Exit Sub
    End If

    tgt.Format.Line.Visible = msoFalse
    Exit Sub

CleanFail:
    MsgError "RemoveOutline"
End Sub

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
Private Function GetLineTarget() As Object
    If Not ActiveChart Is Nothing Then
        If Not Selection Is Nothing Then
            If HasLineFormat(Selection) Then
                Set GetLineTarget = Selection
                Exit Function
            End If
        End If
        ' fallback
        Set GetLineTarget = ActiveChart.ChartArea
        Exit Function
    End If

    ' shapes or ranges
    If Not Selection Is Nothing Then
        If HasLineFormat(Selection) Then Set GetLineTarget = Selection
    End If
End Function


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


Private Function HasLineFormat(o As Object) As Boolean
    On Error Resume Next
    Dim x: Set x = o.Format.Line
    HasLineFormat = (Err.Number = 0)
    Err.Clear
End Function

Private Function HasFillFormat(o As Object) As Boolean
    On Error Resume Next
    Dim x: Set x = o.Format.Fill
    HasFillFormat = (Err.Number = 0)
    Err.Clear
End Function
