Attribute VB_Name = "modFormatFill"
'==== Module: modFormatFill ====
Option Explicit

'   TAG DISPATCHER
' Called from modRibbonHandlers. Parses the ribbon button tag and calls ApplyFill or RemoveFill.
' Tag format: "FILL:ColorName" | "FILL:ColorName|0.3" | "FILL:NONE"
Public Sub ApplyFillFromTag(ByVal tagValue As String)
    tagValue = Trim$(tagValue)

    If InStr(1, tagValue, ":", vbTextCompare) = 0 Then
        MsgBox "Invalid Tag. Expected 'Fill:Color|t'.", vbExclamation
        Exit Sub
    End If

    Dim parts() As String
    parts = Split(tagValue, ":")

    Dim payload As String: payload = UCase$(parts(1))

    If payload = "NONE" Or payload = "NOFILL" Or payload = "OFF" Then
        RemoveFill
        Exit Sub
    End If

    Dim subp() As String
    subp = Split(payload, "|")

    Dim colorName As String: colorName = subp(0)
    Dim transparency As Double: transparency = 0

    If UBound(subp) >= 1 Then
        If IsNumeric(subp(1)) Then transparency = CDbl(subp(1))
    End If

    Dim colorRGB As Long: colorRGB = ColorFromName(colorName)
    If colorRGB = -1 Then
        MsgBox "Unknown color '" & colorName & "'", vbExclamation
        Exit Sub
    End If

    ApplyFill colorRGB, transparency
End Sub


Private Function ColorFromName(ByVal name As String) As Long
    Select Case UCase$(name)
        Case "OCEAN":    ColorFromName = colorOcean
        Case "CORAL":    ColorFromName = colorCoral
        Case "SKY":      ColorFromName = colorSky
        Case "PINE":     ColorFromName = colorPine
        Case "GOLD":     ColorFromName = colorGold
        Case "RUST":     ColorFromName = colorRust
        Case "LAVENDER": ColorFromName = colorLavender
        Case "SILVER":   ColorFromName = colorSilver
        Case "WHITE":    ColorFromName = colorWhite
        Case Else:       ColorFromName = -1
    End Select
End Function


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
    ' Resolve a chart even when a ribbon button click deactivated it before
    ' onAction fired — in that case ActiveChart is Nothing but the ChartObject
    ' remains selected at the worksheet level.
    Dim cht As Chart
    If Not ActiveChart Is Nothing Then
        Set cht = ActiveChart
    ElseIf TypeName(Selection) = "ChartObject" Then
        Set cht = Selection.Chart
    End If

    If Not cht Is Nothing Then
        If Not Selection Is Nothing Then
            If HasFillFormat(Selection) And TypeName(Selection) <> "ChartObject" Then
                Set GetFillTarget = Selection
                Exit Function
            End If
        End If
        Set GetFillTarget = cht.ChartArea
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
