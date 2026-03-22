Attribute VB_Name = "modFormatFill"
'==== Module: modFormatFill ====
Option Explicit

'   TAG DISPATCHER
' Called from modRibbonHandlers. Parses the ribbon button tag and calls ApplyFill or RemoveFill.
' Tag format: "FILL:ColorName" | "FILL:ColorName|0.3" | "FILL:NONE"
Public Sub ApplyFillFromTag(ByVal tagValue As String)
    tagValue = Trim$(tagValue)

    If InStr(1, tagValue, ":", vbTextCompare) = 0 Then
        MsgInvalidFillTag
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
        MsgUnknownColor colorName
        Exit Sub
    End If

    ApplyFill colorRGB, transparency
End Sub


Private Function ColorFromName(ByVal name As String) As Long
    Select Case UCase$(name)
        Case "DATA1":    ColorFromName = colorData1
        Case "DATA2":    ColorFromName = colorData2
        Case "DATA3":    ColorFromName = colorData3
        Case "DATA4":    ColorFromName = colorData4
        Case "DATA5":    ColorFromName = colorData5
        Case "DATA6":    ColorFromName = colorData6
        Case "DATA7":    ColorFromName = colorData7
        Case "NEUTRAL1": ColorFromName = colorNeutral1
        Case "NEUTRAL4": ColorFromName = colorNeutral4
        Case Else:       ColorFromName = -1
    End Select
End Function


'   FILL
Public Sub ApplyFill(ByVal colorRGB As Long, Optional ByVal transparency As Single = 0!)
    ' transparency: 0 = opaque, 1 = fully transparent (chart model)
    ' For line/scatter series the colour lives on .Format.Line; transparency is not applicable.
    On Error GoTo CleanFail

    Dim tgt As Object
    Set tgt = GetFillTarget()
    If tgt Is Nothing Then
        MsgSelectTarget
        Exit Sub
    End If

    If IsLineTarget(tgt) Then
        With tgt.Format.Line
            .Visible = msoTrue
            .ForeColor.rgb = colorRGB
        End With
    Else
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
    End If

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

    If IsLineTarget(tgt) Then
        tgt.Format.Line.Visible = msoFalse
    Else
        tgt.Format.Fill.Visible = msoFalse
    End If

    Exit Sub

CleanFail:
    MsgError "RemoveFill"
End Sub

'   TARGET DETECTION HELPERS

Private Function IsLineTarget(ByVal tgt As Object) As Boolean
    ' Returns True if tgt is a series whose chart type is line or scatter.
    ' Used to decide whether to write to .Format.Line rather than .Format.Fill.
    Dim srs As Series
    On Error Resume Next
    Set srs = tgt
    On Error GoTo 0
    If srs Is Nothing Then Exit Function    ' not a series — use FILL

    Dim ct As Long
    On Error Resume Next
    ct = srs.ChartType
    On Error GoTo 0

    Select Case ct
        Case xlLine, xlLineMarkers, xlLineStacked, xlLineMarkersStacked, _
             xlLineStacked100, xlLineMarkersStacked100, _
             xlXYScatter, xlXYScatterLines, xlXYScatterLinesNoMarkers, _
             xlXYScatterSmooth, xlXYScatterSmoothNoMarkers
            IsLineTarget = True
    End Select
End Function


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
