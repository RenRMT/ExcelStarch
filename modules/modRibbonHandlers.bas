Attribute VB_Name = "modRibbonHandlers"
'==== Module: modRibbonHandlers ====
' All ribbon onAction callbacks
' Each handler is a thin wrapper calling the public entry point in its implementation module.
' Mirrors the ribbon XML structure: one file to audit all button-to-handler mappings.
Option Explicit

'=== Chart creation ===
Public Sub Bar_onAction(control As IRibbonControl): BarChart: End Sub
Public Sub BarwithBlueRamp_onAction(control As IRibbonControl): BarChartBlueRamp: End Sub

Public Sub Column_onAction(control As IRibbonControl): ColumnChart: End Sub
Public Sub ColumnwithBlueRamp_onAction(control As IRibbonControl): ColumnChartBlueRamp: End Sub

Public Sub StackedBar_onAction(control As IRibbonControl): StackedBarChart: End Sub
Public Sub StackedBarwithBlueRamp_onAction(control As IRibbonControl): StackedBarChartBlueRamp: End Sub

Public Sub StackedColumn_onAction(control As IRibbonControl): StackedColumnChart: End Sub
Public Sub StackedColumnwithBlueRamp_onAction(control As IRibbonControl): StackedColumnChartBlueRamp: End Sub

Public Sub Area_onAction(control As IRibbonControl): AreaChart: End Sub
Public Sub AreawithBlueRamp_onAction(control As IRibbonControl): AreaChartBlueRamp: End Sub

Public Sub LineChart_onAction(control As IRibbonControl): LineChart: End Sub
Public Sub LinewithMarkers_onAction(control As IRibbonControl): MarkersLineChart: End Sub

Public Sub Scatter_onAction(control As IRibbonControl): ScatterChart: End Sub
Public Sub Pie_onAction(control As IRibbonControl): PieChart: End Sub
Public Sub StyleSlopeChart_onAction(control As IRibbonControl): SlopeChart: End Sub
Public Sub StyleDotPlot_onAction(control As IRibbonControl): DotPlot: End Sub

'=== Chart tools ===
Public Sub RemoveLegendResizeButton_onAction(control As IRibbonControl): RemoveLegendResizeButton: End Sub
Public Sub LabelLastPointButton_onAction(control As IRibbonControl): LabelLastPointButton: End Sub

'=== Format (fill / outline) ===
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
    '   PARSE e.g. "OCEAN|0.3"
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
        ApplyFill rgb, arg                  ' arg = transparency 0-1
    Else
        MsgBox "Unknown mode '" & mode & "'", vbExclamation
    End If
End Sub

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

'=== Utilities ===
Public Sub StartWithGrayButton_onAction(control As IRibbonControl): StartWithGray: End Sub
Public Sub ChartExport_onAction(control As IRibbonControl): ExportChart: End Sub
Public Sub NotesButton_onAction(control As IRibbonControl): NotesButton: End Sub
