Attribute VB_Name = "modConfig"
Option Explicit
'==== Module: modConfig ====
'
' +---------------------------------------------------------+
' |  SECTION 1 USER SETTINGS                                |
' |  Edit these constants to customise the chart style.     |
' +---------------------------------------------------------+

'=== Identity ===
Public Const orgName As String = "COMPANY"

'=== Canvas ===
Public Const chartWidth As Double = 600         ' 20cm canvas width
Public Const chartHeight As Double = 600        ' 20cm canvas height

'=== Fonts ===
Public Const fontPrimary As String = "Calibri"
Public Const fontPrimaryItalic As String = "Calibri Italic"

' Font sizes
Public Const titleFontSize As Double = 28
Public Const subTitleFontSize As Double = 22
Public Const generalFontSize As Double = 18
Public Const axisFontSize As Double = 18
Public Const sourceTextFontSize As Double = 14

'Font colors
Public Const titleFontColor As Long = colorBrand1
Public Const subTitleFontColor As Long = colorBrand2
Public Const axisFontColor As Long = colorBrand3
Public Const legendFontColor As Long = colorBrand3
Public Const generalFontColor As Long = colorBrand3
Public Const sourceFontColor As Long = colorBrand3
Public Const figureFontColor As Long = colorBrand3

'=== Data series ===
Public Const seriesGapWidth As Double = 33
Public Const seriesOverlap As Double = -5
Public Const lollipopGapWidth As Double = 150     ' wider gap for cleaner stem spacing
Public Const lollipopStickWeight As Single = 2 ' error bar line weight in points

'=== Layout — label last point ===
Public Const labelLastPointPlotWidthInset As Long = 66    ' narrowed for end labels on line charts
Public Const labelLastPointPlotTop As Long = 105
Public Const labelLastPointPlotWidthRatio As Double = 0.98
Public Const labelLastPointTitleNudge As Long = -13

'=== Layout — pie ===
Public Const piePlotAreaSize_legend As Long = 421   ' width and height (square) when legend present
Public Const piePlotAreaSize_noLegend As Long = 447 ' width and height (square) without legend
Public Const piePlotAreaLeft As Long = 131
Public Const piePlotAreaTop As Long = 53
Public Const piePlotTopRatio As Double = 0.75   ' vertical centering ratio
Public Const pieLegendTop As Long = 79

'=== Layout — error textbox ===
Public Const errorBoxWidth As Long = 657
Public Const errorBoxHeight As Long = 53
Public Const errorBoxFontSize As Double = 10


'=== Weights ===
Public Const gridlineWeight As Double = 1
Public Const axisLineWeight As Double = 1

'=== Default placeholder texts ===
Public Const figureBoxDefaultText As String = "Figure XX (optional)"
Public Const titleDefaultText    As String = "Title in 28pt sentence case"
Public Const subtitleDefaultText As String = "Subtitle in 22pt sentence case"
Public Const yAxisDefaultText    As String = "Y axis title (unit)"
Public Const xAxisDefaultText    As String = "X axis title (unit)"
Public Const sourceDefaultText   As String = "Source: Source text goes here."
Public Const notesDefaultText    As String = "Notes: Notes text goes here."

'=== Export ===
Public Const exportSection As String = "Chart Export"
Public Const exportSettingKey As String = "File Filter"
Public Const exportDefaultExt As String = "png"
Public Const exportDefaultName As String = "MyChart"


' === Box sizes ===
' Box sizes expressed as proportion of chart width / height
'top
Public Const FigureBoxHeightProportion As Double = 0.04
Public Const titleBoxHeightProportion As Double = 0.06
Public Const subtitleBoxHeightProportion As Double = 0.05
Public Const yAxisLabelHeightProportion As Double = 0.04
Public Const legendHeightProportion As Double = 0.04
Public Const titleBoxWidthProportion As Double = 1 ' keep this as one unless you are moving logo to the top
Public Const titleBoxNudgeProportion As Double = 0
'bottom
Public Const sourceBoxWidthProportion As Double = 0.8
Public Const sourceBoxHeightProportion As Double = 0.08
Public Const sourceBoxNudgeProportion As Double = 0.01

'padding
Public Const legendLeftPadProportion As Double = 0.01
Public Const plotAreaLeftProportion As Double = 0.005
Public Const xAxisTitle_plotGap As Double = 20        ' gap between plot area base and x-axis title

'=== Layout and logo ===
Public Const logoHeightScale As Double = 0.1        ' logo height as fraction of chart height
Public Const logoAspectRatio As Double = 1          ' logo width = aspectRatio x height
Public Const logoMarginRight As Double = chartWidth * 0.01
Public Const logoMarginBottom As Double = chartHeight * 0.01


' +---------------------------------------------------------+
' |  DERIVED CONSTANTS                                      |
' |  Computed from Section 1. Do not edit directly.         |
' +---------------------------------------------------------+
'=== Logo geometry ===
Public Const logoHeight As Double = chartHeight * logoHeightScale
Public Const logoTop As Double = chartHeight - logoHeight - logoMarginBottom

'=== Title area ===
Public Const figureBoxTop As Double = 0

Public Const figureBoxHeight As Double = chartHeight * FigureBoxHeightProportion
Public Const titleBoxHeight As Double = 30 'chartHeight * titleBoxHeightProportion
Public Const subtitleBoxHeight As Double = chartHeight * subtitleBoxHeightProportion
Public Const titleBoxWidth As Double = chartWidth * titleBoxWidthProportion
Public Const titleBoxNudge As Double = chartWidth * titleBoxNudgeProportion
'Calculations
Public Const calcTitlesHeight As Double = figureBoxHeight + titleBoxHeight + subtitleBoxHeight

'=== Plot area ===
' All values are in Excel points (1pt = 1/72" or about 13/360cm). Origin is the top-left corner of the
' chart area. Changing plotAreaHeight or plotAreaTop will shift the x-axis title, and
' may require adjusting legendTop and the yAxisLabel_* constants to keep elements
' visually balanced. All plot area bottoms (Top + Height) must stay below logoTop (532 pt).
Public Const legendTop As Double = calcTitlesHeight
Public Const legendLeftPad As Double = chartWidth * legendLeftPadProportion
Public Const LegendHeight As Double = chartHeight * legendHeightProportion

'Y axis
Public Const yAxisLabelTop As Double = calcTitlesHeight + LegendHeight
Public Const yAXisLabelHeight As Double = chartHeight * yAxisLabelHeightProportion
Public Const yAxisLabelTop_noLegend As Double = calcTitlesHeight

' Plot area
Public Const plotAreaWidth As Double = chartWidth
Public Const plotAreaLeft As Double = chartWidth * plotAreaLeftProportion

Public Const PlotAreaTop As Double = yAxisLabelTop + yAXisLabelHeight
Public Const PlotAreaTop_noLegend As Double = yAxisLabelTop_noLegend + yAXisLabelHeight

Public Const PlotAreaHeight_noLegend As Double = chartHeight - calcTitlesHeight - yAXisLabelHeight - logoHeight
Public Const PlotAreaHeight As Double = PlotAreaHeight_noLegend - LegendHeight

'Calculations

'=== Source box ===
Public Const sourceBoxWidth As Double = chartWidth * sourceBoxWidthProportion
Public Const sourceBoxLeftNudge As Double = chartWidth * sourceBoxNudgeProportion
Public Const sourceBoxHeight As Double = chartHeight * sourceBoxHeightProportion

'=== Title area ===

Public Const titleBoxTop As Double = figureBoxHeight
Public Const subtitleBoxTop As Double = titleBoxTop + titleBoxHeight


'=== Remove legend resize — mirrors noLegend-multi plot area ===
Public Const removelegendHeight As Double = PlotAreaHeight_noLegend
Public Const removelegendTop As Double = PlotAreaTop_noLegend
Public Const removeLegend_Width As Double = plotAreaWidth
Public Const removeLegend_Left As Double = plotAreaLeft

'=== Export ===
Public Const exportAppName As String = orgName & " Chart Styles"
