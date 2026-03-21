Attribute VB_Name = "modRamp"
'==== Module: modRamp ====
' Applies colour ramps to the data series of the active chart.
'
' ApplyColorRamp      — single-hue ramp, steps assigned in spread order 5,1,3,6,2,4,7.
' InvertColorRamp     — reverses the current fill colour assignment across all series.
' ApplyDivergingRamp  — two-hue diverging ramp: dark→light on the left, light→dark on
'                       the right, with an optional grey centre for odd series counts.
'
' Step selection for both single and diverging ramps follows the same priority sequence
' [5,1,3,6,2,4,7]. For diverging ramps the selected steps are then sorted numerically
' (1 = lightest, 7 = darkest) before being assigned as a gradient.
'
' Maximum series: 7 (single), 15 (diverging: 7 + grey + 7).
Option Explicit

' ============================================================
'   PUBLIC ENTRY POINTS
' ============================================================

Public Sub InvertColorRamp()
    Dim cht As Chart
    Set cht = ResolveActiveChart()
    If cht Is Nothing Then
        MsgNoActiveChart
        Exit Sub
    End If

    Dim n As Long
    n = cht.SeriesCollection.Count
    If n < 2 Then Exit Sub

    ' Snapshot current fill colors
    Dim colors() As Long
    ReDim colors(1 To n)

    Dim i As Long
    For i = 1 To n
        colors(i) = cht.SeriesCollection(i).Format.Fill.ForeColor.rgb
    Next i

    ' Re-apply in reverse order
    For i = 1 To n
        With cht.SeriesCollection(i).Format.Fill
            .Visible = msoTrue
            .Solid
            .ForeColor.rgb = colors(n - i + 1)
        End With
    Next i
End Sub

Public Sub ApplyColorRamp(ByVal rampName As String)
    Dim cht As Chart
    Set cht = ResolveActiveChart()
    If cht Is Nothing Then
        MsgNoActiveChart
        Exit Sub
    End If
    BuildColorRamp cht, UCase$(Trim$(rampName))
End Sub

Public Sub ApplyDivergingRampFromTag(ByVal tagValue As String)
    Dim parts() As String
    parts = Split(UCase$(Trim$(tagValue)), "|")
    If UBound(parts) < 1 Then
        MsgInvalidDivergingTag
        Exit Sub
    End If
    ApplyDivergingRamp parts(0), parts(1)
End Sub

Public Sub ApplyDivergingRamp(ByVal leftRamp As String, ByVal rightRamp As String)
    Dim cht As Chart
    Set cht = ResolveActiveChart()
    If cht Is Nothing Then
        MsgNoActiveChart
        Exit Sub
    End If
    BuildDivergingRamp cht, UCase$(Trim$(leftRamp)), UCase$(Trim$(rightRamp))
End Sub


' ============================================================
'   PRIVATE BUILDERS
' ============================================================

Private Sub BuildColorRamp(cht As Chart, ByVal rampName As String)
    Dim n As Long
    n = cht.SeriesCollection.Count
    If n = 0 Then Exit Sub

    If n > 7 Then
        MsgRampTooManySeries
        Exit Sub
    End If

    Dim palette(1 To 7) As Long
    If Not LoadPalette(rampName, palette) Then Exit Sub

    ' Assign ramp steps in spread order: 5, 1, 3, 6, 2, 4, 7
    Dim order(1 To 7) As Integer
    order(1) = 5: order(2) = 1: order(3) = 3: order(4) = 6
    order(5) = 2: order(6) = 4: order(7) = 7

    Dim i As Long
    For i = 1 To n
        With cht.SeriesCollection(i).Format.Fill
            .Visible = msoTrue
            .Solid
            .ForeColor.rgb = palette(order(i))
        End With
    Next i
End Sub

Private Sub BuildDivergingRamp(cht As Chart, ByVal leftRamp As String, ByVal rightRamp As String)
    Dim n As Long
    n = cht.SeriesCollection.Count
    If n = 0 Then Exit Sub

    If n > 15 Then
        MsgDivergingTooManySeries
        Exit Sub
    End If

    Dim leftPalette(1 To 7) As Long
    Dim rightPalette(1 To 7) As Long
    If Not LoadPalette(leftRamp, leftPalette) Then Exit Sub
    If Not LoadPalette(rightRamp, rightPalette) Then Exit Sub

    Dim sideCount As Long
    sideCount = n \ 2                       ' floor(N/2)
    Dim hasMiddle As Boolean
    hasMiddle = (n Mod 2 = 1)

    ' Pick first sideCount steps from priority order, then sort ascending (1=lightest)
    Dim priority(1 To 7) As Integer
    priority(1) = 5: priority(2) = 1: priority(3) = 3: priority(4) = 6
    priority(5) = 2: priority(6) = 4: priority(7) = 7

    Dim steps() As Integer
    ReDim steps(1 To sideCount)
    Dim i As Integer, j As Integer, tmp As Integer

    For i = 1 To sideCount
        steps(i) = priority(i)
    Next i

    ' Bubble sort steps ascending (lightest → darkest)
    For i = 1 To sideCount - 1
        For j = 1 To sideCount - i
            If steps(j) > steps(j + 1) Then
                tmp = steps(j): steps(j) = steps(j + 1): steps(j + 1) = tmp
            End If
        Next j
    Next i

    ' Left side: descending through sorted steps (dark → light)
    Dim seriesIdx As Long
    seriesIdx = 1
    For i = sideCount To 1 Step -1
        With cht.SeriesCollection(seriesIdx).Format.Fill
            .Visible = msoTrue
            .Solid
            .ForeColor.rgb = leftPalette(steps(i))
        End With
        seriesIdx = seriesIdx + 1
    Next i

    ' Centre: grey if odd series count
    If hasMiddle Then
        With cht.SeriesCollection(seriesIdx).Format.Fill
            .Visible = msoTrue
            .Solid
            .ForeColor.rgb = colorBrand4
        End With
        seriesIdx = seriesIdx + 1
    End If

    ' Right side: ascending through sorted steps (light → dark)
    For i = 1 To sideCount
        With cht.SeriesCollection(seriesIdx).Format.Fill
            .Visible = msoTrue
            .Solid
            .ForeColor.rgb = rightPalette(steps(i))
        End With
        seriesIdx = seriesIdx + 1
    Next i
End Sub


' ============================================================
'   SHARED HELPERS
' ============================================================

' Returns ActiveChart if a chart is in edit mode, or the chart from a selected
' ChartObject (single-click). Handles ribbon buttons deactivating the chart before
' the onAction callback fires.
Private Function ResolveActiveChart() As Chart
    If Not ActiveChart Is Nothing Then
        Set ResolveActiveChart = ActiveChart
    ElseIf TypeName(Selection) = "ChartObject" Then
        Set ResolveActiveChart = Selection.Chart
    End If
End Function

' Fills a 1-to-7 Long array with the ramp constants for rampName.
' Returns False and shows an error if the name is unrecognised.
Private Function LoadPalette(ByVal rampName As String, palette() As Long) As Boolean
    Select Case rampName
        Case "A"
            palette(1) = rampA1: palette(2) = rampA2: palette(3) = rampA3
            palette(4) = rampA4: palette(5) = rampA5: palette(6) = rampA6
            palette(7) = rampA7
        Case "B"
            palette(1) = rampB1: palette(2) = rampB2: palette(3) = rampB3
            palette(4) = rampB4: palette(5) = rampB5: palette(6) = rampB6
            palette(7) = rampB7
        Case "C"
            palette(1) = rampC1: palette(2) = rampC2: palette(3) = rampC3
            palette(4) = rampC4: palette(5) = rampC5: palette(6) = rampC6
            palette(7) = rampC7
        Case "D"
            palette(1) = rampD1: palette(2) = rampD2: palette(3) = rampD3
            palette(4) = rampD4: palette(5) = rampD5: palette(6) = rampD6
            palette(7) = rampD7
        Case "E"
            palette(1) = rampE1: palette(2) = rampE2: palette(3) = rampE3
            palette(4) = rampE4: palette(5) = rampE5: palette(6) = rampE6
            palette(7) = rampE7
        Case "F"
            palette(1) = rampF1: palette(2) = rampF2: palette(3) = rampF3
            palette(4) = rampF4: palette(5) = rampF5: palette(6) = rampF6
            palette(7) = rampF7
        Case "G"
            palette(1) = rampG1: palette(2) = rampG2: palette(3) = rampG3
            palette(4) = rampG4: palette(5) = rampG5: palette(6) = rampG6
            palette(7) = rampG7
        Case Else
            MsgUnknownRamp rampName
            LoadPalette = False
            Exit Function
    End Select
    LoadPalette = True
End Function
