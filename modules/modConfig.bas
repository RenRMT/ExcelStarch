Attribute VB_Name = "modConfig"
'==== Module: modConfig ====
'=== Chart size constants ===
Public Const chartWidth As Double = 456.48 '6.34" (456.58px) for web
Public Const chartHeight As Double = 456.48 '6.34" (456.58px) for web

'=== Font constants ===
Public Const fontPrimary As String = "Calibri"
Public Const fontPrimaryItalic As String = "Calibri Italic"


'== Organisation identity == 
Public Const orgName As String = "COMPANY"
Public Const orgSupportContact As String = "COMPANY IT"

'data series constants
Public Const seriesGapWidth As Double = 33
Public Const seriesOverlap As Double = 0

'chart area constants
Public Const titleFontSize As Double = 20
Public Const subTitleFontSize As Double = 16
Public Const axisFontSize As Double = 12

'=== modChartBuilder: OuterFormat ===
' Layout constants use Excel points as the unit (1pt = 1/72"). Origin is the top-left
' corner of the chart area. All values are interdependent: changing plotAreaHeight or
' plotAreaTop will shift the x-axis title, and may require adjusting legend_top and the
' yAxisLabel_* constants to keep elements visually balanced. gdChartWidth/Height_web
' define the overall canvas; plot area dimensions must fit within those bounds.
Public Const plotAreaHeight As Long = 310
Public Const plotAreaWidth As Long = 740
Public Const plotAreaTop_default As Long = 120      ' single-series, or multi-series with legend
Public Const plotAreaTop_noLegend As Long = 88      ' multi-series, no legend
Public Const plotAreaLeft As Long = 2

'=== modChartBuilder: FormatXAxisTitle ===
Public Const xAxisTitle_legendGap As Long = 15      ' gap between plot area and x-axis title
Public Const legend_top As Long = 70
Public Const legend_leftPad As Long = 5
Public Const plotArea_noLegendSingleHeight As Long = 320
Public Const plotArea_noLegendSingleTop As Long = 80
Public Const plotArea_noLegendMultiHeight As Long = 350
Public Const plotArea_noLegendMultiTop As Long = 60

'=== modChartBuilder: InsertLogo ===
Public Const logoHeightScale As Double = 0.1        ' logo height as fraction of chart height
Public Const logoAspectRatio As Double = 1.8        ' logo width = aspectRatio x height
Public Const logoMarginRight As Single = 10
Public Const logoMarginBottom As Single = 8

'=== modChartBuilder: InsertSource ===
Public Const sourceBoxWidth As Long = 175
Public Const sourceBoxHeight As Long = 35
Public Const sourceTextFontSize As Double = 11
Public Const sourceBoxLeftNudge As Long = 4
' Source box text is set bold then un-bolded selectively, keeping "Source" and "Notes"
' labels bold and the colon + body text plain. Offsets into the combined string:
'   "Source: Source text goes here.\nNotes: Notes text goes here."
'    123456 7                      3132    37
Public Const sourceBox_sourceUnboldStart As Long = 7    ' ":" after "Source"
Public Const sourceBox_sourceUnboldLen As Long = 24     ' ": Source text goes here."
Public Const sourceBox_notesUnboldStart As Long = 37    ' ":" after "Notes"
Public Const sourceBox_notesUnboldLen As Long = 22      ' ": Notes text goes here."

'=== modChartBuilder: FormatTitle ===
Public Const titleBoxWidth As Long = 250
Public Const titleBoxHeight As Long = 30
Public Const subtitleBoxTop As Long = 40
Public Const subtitleBoxHeight As Long = 25
Public Const titleBoxNudge As Long = 4              ' pixel nudge applied to top/left for alignment
Public Const yAxisLabel_legendTop As Long = 96
Public Const yAxisLabel_singleTop As Long = 65
Public Const yAxisLabel_multiTop As Long = 52
Public Const yAxisLabel_legendHeight As Long = 20
Public Const yAxisLabel_noLegendHeight As Long = 18

'=== modChartBuilder: FormatGridlines ===
Public Const gridlineWeight As Double = 1

'=== modChartBuilder: FormatXAxis ===
Public Const axisLineWeight As Double = 1

'=== modChartPie: SetPieChartSizeandTitle ===
Public Const pieTitleFontSize As Double = 18
Public Const pieTitleBoxHeight As Long = 25         ' pie title box height (smaller than standard)
Public Const pieSubtitleFontSize As Double = 14
Public Const pieSubtitleBoxTop As Long = 25
Public Const pieSubtitleBoxHeight As Long = 20
Public Const pieYAxisLabelBoxTop As Long = 45
Public Const piePlotAreaSize_legend As Long = 320   ' width and height (square) when legend present
Public Const piePlotAreaSize_noLegend As Long = 340 ' width and height (square) without legend
Public Const piePlotAreaLeft_web As Long = 100
Public Const piePlotAreaTop_web As Long = 40
Public Const piePlotTopRatio_web As Double = 0.75   ' vertical centering ratio
Public Const pieLegendTop_web As Long = 60

'=== error textbox (modChartPie, modChartScatter, modFormatSeries) ===
Public Const errorBoxWidth As Long = 500
Public Const errorBoxHeight As Long = 40
Public Const errorBoxFontSize As Double = 10

'=== modRemoveLegendResize ===
Public Const removeLegend_webHeight As Long = 170
Public Const removeLegend_webTop As Long = 65
Public Const removeLegend_webWidth As Long = 300
Public Const removeLegend_webLeft As Long = 1

'=== modLabelLastPoint: BuildLabelLastPoint ===
Public Const labelLastPointPlotWidthInset As Long = 50    ' narrowed for end labels on line charts
Public Const labelLastPointPlotTop As Long = 80
Public Const labelLastPointPlotWidthRatio_web As Double = 0.98
Public Const labelLastPointTitleNudge As Long = -10

'=== modExport ===
Public Const exportAppName As String = orgName & " Chart Styles"
Public Const exportSection As String = "Chart Export"
Public Const exportSettingKey As String = "File Filter"
Public Const exportDefaultExt As String = "png"
Public Const exportDefaultName As String = "MyChart"
