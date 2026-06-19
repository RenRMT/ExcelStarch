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

' Generic error handler - call from a CleanFail label while Err object is populated
Public Sub MsgError(ByVal source As String)
    MsgBox source & ": " & Err.Number & " - " & Err.Description, vbExclamation
End Sub

' Guard: too many series for a colour ramp (max 10)
Public Sub MsgRampTooManySeries()
    MsgBox "Colour ramps support a maximum of 10 data series.", vbExclamation, "Too Many Series"
End Sub

' Guard: too many series for a diverging colour ramp (max 21: 10 + grey + 10)
Public Sub MsgDivergingTooManySeries()
    MsgBox "Diverging colour ramps support a maximum of 21 data series.", vbExclamation, "Too Many Series"
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

' ToggleLegend: chart does not support a legend (single-series or treemap)
Public Sub MsgLegendNotApplicable()
    MsgBox "This chart has only 1 data series or does not support a legend.", vbInformation, "Legend Not Applicable"
End Sub

' TogglePaletteOrder: confirms new palette state after toggle
Public Sub MsgPaletteOrderToggled(ByVal altOrder As Boolean)
    If altOrder Then
        MsgBox "Palette order: Violet, Baltic, Sky, Teal, Jasmine, Blush, Coral, Cherry (Rainbow)", vbInformation, "Palette Order"
    Else
        MsgBox "Palette order: Teal, Jasmine, Baltic, Coral, Sky, Cherry, Blush, Violet (Contrasting, default)", vbInformation, "Palette Order"
    End If
End Sub

' ApplyFill: no specific series selected when fill color button clicked
Public Sub MsgSelectSeries()
    MsgBox "Select a specific data series or data point to apply fill colour.", vbExclamation, "No Series Selected"
End Sub

' BuildChartExChrome: chartex chrome (title/logo/source) is placed on the host
' worksheet, so the chart must be embedded - chart sheets are unsupported.
Public Sub MsgChartExNeedsEmbedded()
    MsgBox "This chart type requires an embedded chart on a worksheet, not a chart sheet.", vbExclamation, "Not Supported Here"
End Sub

' GroupChartExChrome: the chrome shapes could not be grouped with the chart;
' they remain on the sheet but are not bound together for moving/exporting.
Public Sub MsgChartExGroupFailed()
    MsgBox "Could not group the chart with its title, logo and source boxes." & vbNewLine & _
           "The shapes were created but are not grouped; select and group them manually before exporting.", _
           vbExclamation, "Grouping Failed"
End Sub

