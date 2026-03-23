Attribute VB_Name = "modConfig"
Option Explicit
'==== Module: modConfig ====
'=== Chart size constants ===
Public Const chartWidth As Double = 600         ' 8.33" canvas width
Public Const chartHeight As Double = 600        ' 8.33" canvas height

'=== Font constants ===
Public Const fontPrimary As String = "Calibri"
Public Const fontPrimaryItalic As String = "Calibri Italic"


'== Organisation identity ==
Public Const orgName As String = "COMPANY"

'data series constants
Public Const seriesGapWidth As Double = 33
Public Const seriesOverlap As Double = 0

'chart area constants
Public Const titleFontSize As Double = 24
Public Const subTitleFontSize As Double = 20
Public Const generalFontSize As Double = 16
Public Const axisFontSize As Double = 16
Public Const titleFontColor As Long = colorBrand1
Public Const subTitleFontColor As Long = colorBrand2
Public Const axisFontColor As Long = colorBrand3
Public Const legendFontColor As Long = colorBrand3
Public Const generalFontColor As Long = colorBrand3

'=== modChartBuilder: OuterFormat ===
' Layout constants use Excel points as the unit (1pt = 1/72"). Origin is the top-left
' corner of the chart area. All values are interdependent: changing plotAreaHeight or
' plotAreaTop will shift the x-axis title, and may require adjusting legend_top and the
' yAxisLabel_* constants to keep elements visually balanced. gdChartWidth/Height_web
' define the overall canvas; plot area dimensions must fit within those bounds.
'
' Logo bottom boundary: logoTop = chartHeight - (chartHeight * logoHeightScale) - logoMarginBottom
'   = 600 - 60 - 8 = 532 pt. All plot area bottoms (Top + Height) must stay below this.
Public Const plotAreaHeight As Long = 370           ' 158 + 370 = 528 < 532 (logo top)
Public Const plotAreaWidth As Long = 973
Public Const plotAreaTop_default As Long = 158      ' single-series, or multi-series with legend
Public Const plotAreaTop_noLegend As Long = 116     ' multi-series, no legend
Public Const plotAreaLeft As Long = 3

'=== modChartBuilder: FormatXAxisTitle ===
Public Const xAxisTitle_plotGap As Long = 20        ' gap between plot area base and x-axis title
Public Const legend_top As Long = 92
Public Const legend_leftPad As Long = 7
Public Const plotArea_noLegendSingleHeight As Long = 421  ' 105 + 421 = 526 < 532 (logo top)
Public Const plotArea_noLegendSingleTop As Long = 105
Public Const plotArea_noLegendMultiHeight As Long = 450   ' 79 + 450 = 529 < 532 (logo top)
Public Const plotArea_noLegendMultiTop As Long = 79

'=== modChartBuilder: InsertLogo ===
Public Const logoHeightScale As Double = 0.1        ' logo height as fraction of chart height
Public Const logoAspectRatio As Double = 1          ' logo width = aspectRatio x height
Public Const logoMarginRight As Single = 10
Public Const logoMarginBottom As Single = 8

'=== modChartBuilder: InsertSource ===
Public Const sourceBoxWidth As Long = 230
Public Const sourceBoxHeight As Long = 46
Public Const sourceTextFontSize As Double = 11
Public Const sourceBoxLeftNudge As Long = 5

'=== modChartBuilder: FormatTitle ===
Public Const titleBoxWidth As Long = 394
Public Const titleBoxHeight As Long = 32    ' reduced from 39 to make room for FigureBox
Public Const subtitleBoxTop As Long = 54   ' = figureBoxHeight(22) + titleBoxHeight(32)
Public Const subtitleBoxHeight As Long = 24 ' reduced from 33 to make room for FigureBox
Public Const figureBoxHeight As Long = 22
Public Const figureBoxDefaultText As String = "Figure XX (optional)"

'=== Placeholder texts (modChartBuilder: FormatTitle, FormatXAxisTitle, InsertSource) ===
Public Const titleDefaultText    As String = "Title in 20pt sentence case"
Public Const subtitleDefaultText As String = "Subtitle in 16pt sentence case"
Public Const yAxisDefaultText    As String = "Y axis title (unit)"
Public Const xAxisDefaultText    As String = "X axis title (unit)"
Public Const sourceDefaultText   As String = "Source: Source text goes here."
Public Const notesDefaultText    As String = "Notes: Notes text goes here."
Public Const titleBoxNudge As Long = 5              ' pixel nudge applied to top/left for alignment
Public Const yAxisLabel_legendTop As Long = 126
Public Const yAxisLabel_singleTop As Long = 85
Public Const yAxisLabel_multiTop As Long = 68
Public Const yAxisLabel_legendHeight As Long = 26
Public Const yAxisLabel_noLegendHeight As Long = 24

'=== modChartBuilder: FormatGridlines ===
Public Const gridlineWeight As Double = 1

'=== modChartBuilder: FormatXAxis ===
Public Const axisLineWeight As Double = 1

'=== modChartPie ===
Public Const piePlotAreaSize_legend As Long = 421   ' width and height (square) when legend present
Public Const piePlotAreaSize_noLegend As Long = 447 ' width and height (square) without legend
Public Const piePlotAreaLeft_web As Long = 131
Public Const piePlotAreaTop_web As Long = 53
Public Const piePlotTopRatio_web As Double = 0.75   ' vertical centering ratio
Public Const pieLegendTop_web As Long = 79

'=== error textbox (modChartPie, modChartScatter, modFormatSeries) ===
Public Const errorBoxWidth As Long = 657
Public Const errorBoxHeight As Long = 53
Public Const errorBoxFontSize As Double = 10

'=== modRemoveLegendResize ===
' Aligned with noLegend-multi dimensions so the plot area matches what the pipeline
' would produce for a multi-series chart created without a legend.
Public Const removeLegend_webHeight As Long = 450   ' = plotArea_noLegendMultiHeight
Public Const removeLegend_webTop As Long = 79       ' = plotArea_noLegendMultiTop
Public Const removeLegend_webWidth As Long = 973    ' = plotAreaWidth
Public Const removeLegend_webLeft As Long = 3       ' = plotAreaLeft

'=== modLabelLastPoint: BuildLabelLastPoint ===
Public Const labelLastPointPlotWidthInset As Long = 66    ' narrowed for end labels on line charts
Public Const labelLastPointPlotTop As Long = 105
Public Const labelLastPointPlotWidthRatio_web As Double = 0.98
Public Const labelLastPointTitleNudge As Long = -13

'=== modChartLollipop ===
Public Const lollipopGapWidth As Long = 150     ' wider gap for cleaner stem spacing
Public Const lollipopStickWeight As Single = 1.5 ' error bar line weight in points

'=== modExport ===
Public Const exportAppName As String = orgName & " Chart Styles"
Public Const exportSection As String = "Chart Export"
Public Const exportSettingKey As String = "File Filter"
Public Const exportDefaultExt As String = "png"
Public Const exportDefaultName As String = "MyChart"
