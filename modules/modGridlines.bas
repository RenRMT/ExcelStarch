Attribute VB_Name = "modGridlines"
'==== Module: modGridlines ====
' Toggles major gridlines on the selected chart, rotating through four states:
'   1. None
'   2. Horizontal only  (value / Y axis)
'   3. Vertical only    (category / X axis)
'   4. Both
'
' Shared constants from modConfig: colorSteel, gridlineWeight
Option Explicit

Public Sub ToggleGridlines()
    If ActiveChart Is Nothing Then
        MsgBox "Please select a chart first.", vbExclamation
        Exit Sub
    End If

    Dim cht As Chart
    Set cht = ActiveChart

    ' Read current state from the chart itself
    Dim hasH As Boolean     ' horizontal gridlines (value / Y axis)
    Dim hasV As Boolean     ' vertical gridlines   (category / X axis)

    If cht.HasAxis(xlValue) Then hasH = cht.Axes(xlValue).HasMajorGridlines
    If cht.HasAxis(xlCategory) Then hasV = cht.Axes(xlCategory).HasMajorGridlines

    ' Advance state: None → Horizontal → Vertical → Both → None
    Dim nextH As Boolean
    Dim nextV As Boolean

    If Not hasH And Not hasV Then
        nextH = True:  nextV = False        ' None → Horizontal only
    ElseIf hasH And Not hasV Then
        nextH = False: nextV = True         ' Horizontal → Vertical only
    ElseIf Not hasH And hasV Then
        nextH = True:  nextV = True         ' Vertical → Both
    Else
        nextH = False: nextV = False        ' Both → None
    End If

    ' Apply or remove horizontal gridlines (value / Y axis)
    If cht.HasAxis(xlValue) Then
        If nextH Then
            ApplyAxisGridlines cht.Axes(xlValue)
        Else
            cht.Axes(xlValue).HasMajorGridlines = False
        End If
    End If

    ' Apply or remove vertical gridlines (category / X axis)
    If cht.HasAxis(xlCategory) Then
        If nextV Then
            ApplyAxisGridlines cht.Axes(xlCategory)
        Else
            cht.Axes(xlCategory).HasMajorGridlines = False
        End If
    End If
End Sub

' Enables major gridlines on an axis and applies standard COMPANY formatting.
Private Sub ApplyAxisGridlines(ax As Axis)
    If Not ax.HasMajorGridlines Then ax.HasMajorGridlines = True

    With ax.MajorGridlines.Format.Line
        .Visible = msoTrue
        .weight = gridlineWeight
        .DashStyle = msoLineSolid
        .ForeColor.rgb = colorSteel
    End With
End Sub
