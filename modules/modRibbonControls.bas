Attribute VB_Name = "modRibbonControls"
'==== Module: modRibbonHandlers ====
Option Explicit

' Ribbon handler for all fill/outline color buttons.
' Tag format (set in ribbon XML):
'   "FILL:ColorName"          — solid fill, no transparency
'   "FILL:ColorName|0.3"      — fill with 30% transparency
'   "FILL:NONE"               — remove fill
'   "OUTLINE:ColorName|2"     — outline with weight 2pt
'   "OUTLINE:NONE"            — remove outline
' Valid color names: OCEAN, CORAL, SKY, PINE, GOLD, RUST, LAVENDER, SILVER, WHITE
Public Sub Format_onAction(control As IRibbonControl)
    Dim tagValue As String
    tagValue = Trim$(control.Tag)

    If InStr(1, tagValue, ":", vbTextCompare) = 0 Then
        MsgBox "Invalid Tag. Expected 'Fill:Color|t' or 'Outline:Color|w'.", vbExclamation
        Exit Sub
    End If

    Dim parts() As String
    parts = Split(tagValue, ":")

    Dim mode As String: mode = UCase$(parts(0))   ' "OUTLINE" or "FILL"
    Dim payload As String: payload = UCase$(parts(1))

    '--------------------------------------------
    '   NO OUTLINE / NO FILL
    '--------------------------------------------
    If payload = "NONE" Or payload = "NOFILL" Or payload = "NOOUTLINE" Or payload = "OFF" Then
        If mode = "OUTLINE" Then
            RemoveOutline
        ElseIf mode = "FILL" Then
            RemoveFill
        End If
        Exit Sub
    End If

    '--------------------------------------------
    '   PARSE e.g. "YELLOW|0.3"
    '--------------------------------------------
    Dim subp() As String
    subp = Split(payload, "|")

    Dim colorName As String: colorName = subp(0)
    Dim arg As Double: arg = 0

    If UBound(subp) >= 1 Then
        If IsNumeric(subp(1)) Then arg = CDbl(subp(1))
    End If

    Dim rgb As Long: rgb = ColorFromName(colorName)
    If rgb = -1 Then
        MsgBox "Unknown color '" & colorName & "'", vbExclamation
        Exit Sub
    End If

    '--------------------------------------------
    '   EXECUTE
    '--------------------------------------------
    If mode = "OUTLINE" Then
        If arg = 0 Then arg = 2            ' default weight when tag omits |weight
        ApplyOutline rgb, arg               ' arg = weight
    ElseIf mode = "FILL" Then
        ApplyFill rgb, arg                  ' arg = transparency 0�1
    Else
        MsgBox "Unknown mode '" & mode & "'", vbExclamation
    End If
End Sub


' Color lookup
Private Function ColorFromName(ByVal name As String) As Long
    Select Case UCase$(name)
        Case "OCEAN": ColorFromName = colorOcean
        Case "CORAL": ColorFromName = colorCoral
        Case "SKY": ColorFromName = colorSky
        Case "PINE": ColorFromName = colorPine
        Case "GOLD": ColorFromName = colorGold
        Case "RUST": ColorFromName = colorRust
        Case "LAVENDER": ColorFromName = colorLavender
        Case "SILVER": ColorFromName = colorSilver
        Case "WHITE": ColorFromName = colorWhite
        Case Else
            ColorFromName = -1
    End Select
End Function

' reset colors to grey button
Public Sub StartWithGrayButton_onAction(control As IRibbonControl)
    StartWithGray
End Sub
