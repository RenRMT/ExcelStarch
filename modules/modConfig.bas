Attribute VB_Name = "modConfig"
Option Explicit
'==== Module: modConfig ====
' This config module consists of two sections:
' Section 1 contains user settings that can be easily adapted to your own preferences.
' These settings include things like chart canvas dimensions, Font types & sizes,
' font colors, logo size, margins, and default title/subtitle texts.
' These can all be safely changed according to your needs.
' Generally, the chart layout should be responsive to any changes you make.
' You might need to test font sizes a little bit.
'
' Section 2 contains calculated & derived constants. These ensure that all
' relevant settings respond correctly to any changes you make in section 1.
' Unless you are planning to significantly deviate from the default chart layout,
' touching these settings is not recommended.


' +---------------------------------------------------------+
' |  SECTION 1 USER SETTINGS                                |
' |  Edit these constants to customise the chart style.     |
' +---------------------------------------------------------+

'=== Identity settings ===
Public Const orgName As String = "COMPANY"

'=== Canvas settings ===
' Canvas is measured in Excel points (1pt = 1/72" or about 13/360cm).
' Origin is the top-left corner of the chart area.
Public Const chartWidth As Double = 600         ' 20cm canvas width
Public Const chartHeight As Double = 600        ' 20cm canvas height

'=== Placeholder text settings ===
' Chart text boxes
' all chart text elements come pre-filled with placeholder text.
' use this text to convey standards relating to chart text,
' and to show the standard font colors for these texts.
Public Const figureBoxDefaultText As String = "Figure XX (optional)"
Public Const titleDefaultText    As String = "Title in 28pt sentence case"
Public Const subtitleDefaultText As String = "Subtitle in 22pt sentence case"
Public Const yAxisDefaultText    As String = "Y axis title (unit)"
Public Const xAxisDefaultText    As String = "X axis title (unit)"
Public Const sourceDefaultText   As String = "Source: Source text goes here."
Public Const notesDefaultText    As String = "Notes: Notes text goes here."

' Export
Public Const exportSection As String = "Chart Export"
Public Const exportSettingKey As String = "File Filter"
Public Const exportDefaultExt As String = "png"
Public Const exportDefaultName As String = "MyChart"


'=== Font settings ===
' Font family
' - fontPrimary: font used for most text boxes, including title, legend, source box.
' - fontPrimaryItalic: by default, italic font is only used for Y-/X-axis labels.
'   Set to same value as fontPrimary if you don't want to use italic font.
Public Const fontPrimary As String = "Calibri"
Public Const fontPrimaryItalic As String = "Calibri Italic"

' Font sizes
' font sizes expressed in Excel points
' - generalFontSize: is currently not used for anything but kept for compatibility
Public Const titleFontSize As Double = 28
Public Const subTitleFontSize As Double = 22
Public Const figureFontSize As Double = 18
Public Const axisFontSize As Double = 18
Public Const sourceTextFontSize As Double = 14
Public Const generalFontSize As Double = 18

'Font colors
' colors as defined in module modConfigColors.
' - generalFontColor: is currently not used for anything but kept for compatibility
Public Const titleFontColor As Long = colorBrand1
Public Const subTitleFontColor As Long = colorBrand2
Public Const figureFontColor As Long = colorBrand3

Public Const axisFontColor As Long = colorBrand3
Public Const legendFontColor As Long = colorBrand3
Public Const sourceFontColor As Long = colorBrand3
Public Const generalFontColor As Long = colorBrand3

'=== Chart Data settings ===
' General data series settings
' - seriesGapWidth: amount of horizontal space between data series, expressed as a
'   percentage of the series width.
' - seriesOverlap: amount of overlap between data series. negative values create distance,
'   positive values create overlap. Setting to 0 makes data series touch. Note that
'   this setting is overriden for stacked bar/column charts.
Public Const seriesGapWidth As Double = 33
Public Const seriesOverlap As Double = -5

' Lollipop chart settings
' Lollipop charts are generated a bit differently and require their own settings.
' - lollipopGapWidth: set width between lollipop chart series.
' - lollipopStickWeight: the size of the lollipop sticks expressed in Excel points
Public Const lollipopGapWidth As Double = 150
Public Const lollipopStickWeight As Single = 2

' Pie chart settings
Public Const pieplotAreaSize_legend As Long = 400   ' width and height (square) when legend present
Public Const pieplotAreaSize_noLegend As Long = 447 ' width and height (square) without legend
Public Const pieplotAreaLeft As Long = 131
Public Const pieplotAreaTop As Long = 53
Public Const piePlotTopRatio As Double = 0.75   ' vertical centering ratio
Public Const pieLegendGap As Double = 6          ' gap between subtitle box and legend
' Note: pieLegendTop is defined in the Title area section below, since it derives
' from subtitleBoxTop/subtitleBoxHeight (VBA Const cannot forward-reference).

' Weights
'   - gridLineWeight:  weight of chart gridlines expressed in Excel points
'   - axisLineWeight:weight of chart axis lines expressed in Excel points
Public Const gridlineWeight As Double = 1
Public Const axisLineWeight As Double = 1

'=== Chart Actions settings ===
'Label last point
Public Const labelLastPointPlotWidthInset As Long = 66    ' narrowed for end labels on line charts
Public Const labelLastPointPlotTop As Long = 105
Public Const labelLastPointPlotWidthRatio As Double = 0.98
Public Const labelLastPointTitleNudge As Long = -13


' === Box sizes ===
' Box sizes expressed as proportion of chart width / height
Public Const FigureBoxHeightProportion As Double = 0.04
Public Const titleBoxHeightProportion As Double = 0.07
Public Const subtitleBoxHeightProportion As Double = 0.05
Public Const yAxisLabelHeightProportion As Double = 0.04
Public Const legendHeightProportion As Double = 0.04
Public Const titleBoxWidthProportion As Double = 1 'keep this as 1 unless you are moving logo to the top
Public Const titleBoxNudgeProportion As Double = 0
'bottom
Public Const sourceBoxWidthProportion As Double = 0.8
Public Const sourceBoxHeightProportion As Double = 0.08
Public Const sourceBoxNudgeProportion As Double = 0.01

'padding
Public Const legendLeftPadProportion As Double = 0
Public Const plotAreaLeftProportion As Double = 0.005
Public Const yAxisLabelPad As Double = 10

'=== Layout and logo ===
' - logoFileType: Embedded logo accepts PNG or SVG files
' - logoHeightScale: as proportion of chart height.
' - logoAspectRatio: the keep aspect ratio setting in Excel does not work properly.
'   setting it here prevents your logo from getting distorted on chart resize.
' - plotAreaBottomMarginProp: reserved space below plot area to prevent X-axis labels
'   from overlapping the logo. Adjust based on axis label height and desired spacing.
Public Const logoFileType As String = "svg"
Public Const logoHeightScale As Double = 0.1        ' logo height as fraction of chart height
Public Const logoAspectRatio As Double = 1          ' logo width = aspectRatio x height
Public Const logoMarginRightProp As Double = 0.01 'chartWidth * 0.01
Public Const logoMarginBottomProp As Double = 0.01 'chartHeight * 0.01
Public Const plotAreaBottomMarginProp As Double = 0.03 ' clearance between x-axis labels and top logo, as fraction of chart height

' +---------------------------------------------------------+
' |  DERIVED CONSTANTS                                      |
' |  Computed from Section 1. Do not edit directly.         |
' +---------------------------------------------------------+
'=== Logo geometry ===
Public Const logoHeight As Double = chartHeight * logoHeightScale
Public Const logoMarginBottom As Double = chartHeight * logoMarginBottomProp
Public Const logoTop As Double = chartHeight - logoHeight - logoMarginBottom
Public Const logoMarginRight As Double = chartWidth * logoMarginRightProp

'=== Plot area bottom margin ===
Public Const plotAreaBottomMargin As Double = chartHeight * plotAreaBottomMarginProp


'=== Title area ===
' Figure number box aligns with the top of the chart area (no canvas margin).
Public Const figureBoxTop As Double = 0
Public Const figureBoxHeight As Double = chartHeight * FigureBoxHeightProportion
Public Const titleBoxTop As Double = figureBoxHeight
Public Const titleBoxHeight As Double = chartHeight * titleBoxHeightProportion
Public Const subtitleBoxTop As Double = titleBoxTop + titleBoxHeight
Public Const subtitleBoxHeight As Double = chartHeight * subtitleBoxHeightProportion
Public Const titleBoxWidth As Double = chartWidth * titleBoxWidthProportion
Public Const titleBoxNudge As Double = chartWidth * titleBoxNudgeProportion
Public Const calcTitlesHeight As Double = figureBoxHeight + titleBoxHeight + subtitleBoxHeight ' For calculation only
' Pie/donut legend sits just below the subtitle box to avoid overlapping it.
Public Const pieLegendTop As Double = subtitleBoxTop + subtitleBoxHeight + pieLegendGap

'=== Plot area ===
'Legend
Public Const legendTop As Double = calcTitlesHeight
Public Const legendLeftPad As Double = chartWidth * legendLeftPadProportion
Public Const LegendHeight As Double = chartHeight * legendHeightProportion
'Y axis
Public Const yAxisLabelTop As Double = calcTitlesHeight + LegendHeight
Public Const yAxisLabelHeight As Double = chartHeight * yAxisLabelHeightProportion
Public Const yAxisLabelTop_noLegend As Double = calcTitlesHeight
'X axis title (mirrors the Y strip; sits in a band just above the logo/source band)
Public Const xAxisLabelHeight As Double = yAxisLabelHeight
Public Const xAxisLabelTop As Double = chartHeight - logoHeight - plotAreaBottomMargin - xAxisLabelHeight
'Plot area
' Plot-area Top/Height are no longer fixed constants: they depend on which title
' bands (y-title strip, x-title strip) and the legend are present, so they are
' computed at runtime by PlotAreaTopFor / PlotAreaHeightFor in modChartBuilder.
' Width/Left are still fixed.
Public Const plotAreaWidth As Double = chartWidth
Public Const plotAreaLeft As Double = chartWidth * plotAreaLeftProportion

'=== Source box ===
Public Const sourceBoxWidth As Double = chartWidth * sourceBoxWidthProportion
Public Const sourceBoxLeftNudge As Double = chartWidth * sourceBoxNudgeProportion
Public Const sourceBoxHeight As Double = chartHeight * sourceBoxHeightProportion

'=== Export ===
Public Const exportAppName As String = orgName & " Chart Styles"
