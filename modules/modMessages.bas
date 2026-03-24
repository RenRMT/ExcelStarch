Attribute VB_Name = "modMessages"
'==== Module: modMessages ====
' Centralised user-facing messages. Call these instead of inline MsgBox.
Option Explicit

' Guard: no chart selected
Public Sub MsgNoActiveChart()
    MsgBox "Select a chart and try again.", vbExclamation, "No Active Chart"
End Sub

' Guard: no valid chart element or shape selected
Public Sub MsgSelectTarget()
    MsgBox "Select a chart element or shape.", vbExclamation, "No Selection"
End Sub

' Generic error handler — call from a CleanFail label while Err object is populated
Public Sub MsgError(ByVal source As String)
    MsgBox source & ": " & Err.Number & " - " & Err.Description, vbExclamation
End Sub

' Guard: too many series for a colour ramp (max 7)
Public Sub MsgRampTooManySeries()
    MsgBox "Colour ramps support a maximum of 7 data series.", vbExclamation, "Too Many Series"
End Sub

' Guard: too many series for a diverging colour ramp (max 15: 7 + grey + 7)
Public Sub MsgDivergingTooManySeries()
    MsgBox "Diverging colour ramps support a maximum of 15 data series.", vbExclamation, "Too Many Series"
End Sub

' Guard: invalid colour mode argument passed to FormatSeriesColors
Public Sub MsgInvalidColorMode()
    MsgBox "Invalid mode. Use ""FILL"" or ""LINE"".", vbExclamation, "FormatSeriesColors"
End Sub

' InsertLogo: logo file could not be decoded from Base64
Public Sub MsgLogoDecodeFailed()
    MsgBox "Failed to decode logo image.", vbExclamation, "Logo Error"
End Sub

' GetTargetChart: user has neither an active chart nor a range selected
Public Sub MsgSelectRangeOrChart()
    MsgBox "Please select a data range or an existing chart.", vbExclamation, "No Selection"
End Sub

' ApplyFillFromTag: ribbon tag string is malformed
Public Sub MsgInvalidFillTag()
    MsgBox "Invalid tag. Expected format: 'Fill:Color' or 'Fill:Color|transparency'.", vbExclamation, "Invalid Tag"
End Sub

' ApplyFillFromTag: colour name in tag is not recognised
Public Sub MsgUnknownColor(ByVal colorName As String)
    MsgBox "Unknown colour '" & colorName & "'.", vbExclamation, "Unknown Colour"
End Sub

' GrayOutChart: confirmation prompt before graying out; returns vbOK or vbCancel
Public Function MsgGrayOutConfirm(ByVal duplicateChart As Boolean) As VbMsgBoxResult
    MsgGrayOutConfirm = MsgBox( _
        IIf(duplicateChart, _
            "This will duplicate your chart and make the duplicate gray.", _
            "This will make your current chart gray."), _
        vbExclamation + vbOKCancel)
End Function

' GrayOutChart: the duplicated chart could not be resolved after Duplicate.Select
Public Sub MsgCouldNotResolveDuplicate()
    MsgBox "Could not resolve duplicated chart.", vbExclamation, "Duplicate Error"
End Sub

' modExport: export on macOS is not supported
Public Sub MsgExportMacUnsupported()
    MsgBox "Chart export is not supported on Mac.", vbExclamation, "Unsupported Platform"
End Sub

' ApplyDivergingRampFromTag: tag string does not contain the expected pipe separator
Public Sub MsgInvalidDivergingTag()
    MsgBox "Invalid diverging ramp tag. Expected format: 'LEFT|RIGHT'.", vbExclamation, "Invalid Tag"
End Sub

' LoadPalette: ramp name in tag is not recognised
Public Sub MsgUnknownRamp(ByVal rampName As String)
    MsgBox "Unknown ramp '" & rampName & "'.", vbExclamation, "Unknown Ramp"
End Sub

' TogglePaletteOrder: confirms new palette state after toggle
Public Sub MsgPaletteOrderToggled(ByVal altOrder As Boolean)
    If altOrder Then
        MsgBox "Palette order: Ocean, Lavender, Sky, Pine, Gold, Coral, Rust", vbInformation, "Palette Order"
    Else
        MsgBox "Palette order: Ocean, Coral, Sky, Pine, Gold, Rust, Lavender (default)", vbInformation, "Palette Order"
    End If
End Sub

' Adds a styled error box to the chart when the series count exceeds the supported limit.
' Overlays a yellow/red warning directly on the chart area.
Public Sub MsgTooManySeries(cht As Chart)
    Dim txtB As Shape
    Set txtB = cht.Shapes.AddTextbox(msoTextOrientationHorizontal, 0, 0, errorBoxWidth, errorBoxHeight)
    With txtB
        .name = "ErrorBox"
        With .TextFrame2.TextRange
            .Text = "You have too many data series for this chart type."
            .Font.Size = errorBoxFontSize
            .Font.name = fontPrimary
            .Font.Fill.ForeColor.RGB = vbRed
            .ParagraphFormat.Alignment = msoTextEffectAlignmentLeft
        End With
        .Fill.ForeColor.RGB = vbYellow
    End With
End Sub
