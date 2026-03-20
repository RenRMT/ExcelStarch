Attribute VB_Name = "_discontinued"
'==============================================================================
' Module: _discontinued
' Holds references to discontinued chart types (Area, MarkersLine, Scatter,
' Slope, DotPlot). The implementation files are retained as _modChart*.bas.
'==============================================================================
Option Explicit

'------------------------------------------------------------------------------
' Ribbon handlers
' Origin: modRibbonHandlers — removed when chart types were discontinued
'------------------------------------------------------------------------------
Public Sub Area_onAction(control As IRibbonControl): AreaChart: End Sub
Public Sub LinewithMarkers_onAction(control As IRibbonControl): MarkersLineChart: End Sub
Public Sub Scatter_onAction(control As IRibbonControl): ScatterChart: End Sub
Public Sub StyleSlopeChart_onAction(control As IRibbonControl): SlopeChart: End Sub
Public Sub StyleDotPlot_onAction(control As IRibbonControl): DotPlot: End Sub

'------------------------------------------------------------------------------
' Configuration constants
' Origin: modConfig — sections annotated for discontinued modules only
'------------------------------------------------------------------------------

' Origin: modConfig === modChartSlope: SlopeChartStyles ===
Public Const slopePlotWidth_web As Long = 700
Public Const slopePlotLeft As Long = 40
Public Const slopePlotTop As Long = 80

' Origin: modConfig === modChartDotPlot / modChartSlope: shared label sizes ===
Public Const dataLabelFontSize_secondary As Double = 11

' Origin: modConfig === modChartDotPlot: BuildDotPlot ===
Public Const dotPlotLabelFontSize As Double = 8
Public Const dotPlotMarkerSize As Long = 6
Public Const dotPlotChartTop As Long = 10
Public Const dotPlotChartLeft As Long = 350

' Origin: modConfig === modChartScatter: BuildScatterChart ===
Public Const scatterMarkerSize As Long = 7
