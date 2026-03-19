Attribute VB_Name = "modConfig"
'==== Module: modConfig ====
Public gWebVersion As Boolean
Public gWebCancel As Boolean
Public gPDFChart As Boolean
Public gPNGChart As Boolean
Public gSVGChart As Boolean
Public gGIFChart As Boolean

'=== Chart size constants ===
Public Const gdChartWidth_print As Double = 450 '6.25"
Public Const gdChartHeight_print As Double = 270 '3.75"
'Public Const gdChartWidth_web As Double = 760.32 '10.56" (760px) for web
Public Const gdChartWidth_web As Double = 456.48 '6.34" (456.58px) for web
Public Const gdChartHeight_web As Double = 456.48 '6.34" (456.58px) for web

'Variables for font checking and application
Public FontStyle As String
Public FontStyleItalic As String
Public FontStyleBlack As String
Public FontStyleBold As String
Public FigureNumberFontStyle As String
Public Const gsPRIMARY_FONT As String = "Calibri"
Public Const gsSECONDARY_FONT As String = "Calibri Light"
Public Const gsDEFAULT_FONT As String = "Arial"
Public Const gsPRIMARY_ITALICS_FONT As String = "Calibri Italic"
Public Const gsPRIMARY_BLACK_FONT As String = "Calibri Black"
Public Const gsPRIMARY_BOLD_FONT As String = "Calibri Bold"

Public Const giRGBbluecolor1 As Long = 15984847 'RGB(207, 232, 243)
Public Const giRGBbluecolor2 As Long = 15520930 'RGB(162, 212, 236)
Public Const giRGBbluecolor3 As Long = 14860147 'RGB(115, 191, 226)
Public Const giRGBbluecolor4 As Long = 14396230 'RGB(70, 171, 219)
Public Const giRGBbluecolor5 As Long = 13800982 '?RGB(22, 150, 210) in Immediate window
Public Const giRGBbluecolor6 As Long = 10383634 'RGB(18, 113, 158)
Public Const giRGBbluecolor7 As Long = 6966282  'RGB(10, 76, 106)

' brand colors
Public Const colorPrimaryBlue As Long = 10963739    'RGB(27, 75, 167)
Public Const colorDarkBlue As Long = 2888711        'RGB(7, 20, 44)
Public Const colorBlack As Long = 655874            'RGB(2, 2, 10)
Public Const colorLightGrey As Long = 16382457      'RGB(249, 249, 249)

' neutral colors
Public Const colorSilver As Long = 13421772     'RGB(204, 204, 204)
Public Const colorSteel As Long = 12303291      'RGB(187, 187, 187)
Public Const colorAsh As Long = 10263708        'RGB(156, 156, 156)
Public Const colorWhite As Long = 16777215      'RGB(255, 255, 255)

' data colors
Public Const colorOcean As Long = 12285696      '?RGB(0, 119, 187)
Public Const colorCoral As Long = 6719743       'RGB(255, 136, 102)
Public Const colorSky As Long = 16764023        'RGB(119, 204, 255)
Public Const colorPine As Long = 8952064        'RGB(0, 153, 136)
Public Const colorGold As Long = 3399167        'RGB(255, 221, 51)
Public Const colorRust As Long = 17578          'RGB(170, 68, 0)
Public Const colorLavender As Long = 15636906   'RGB(170, 153, 238)


Public Const rampOcean1 As Long = 15984847 'RGB(207, 232, 243)
Public Const rampOcean2 As Long = 15520930 'RGB(162, 212, 236)
Public Const rampOcean3 As Long = 14860147 'RGB(115, 191, 226)
Public Const rampOcean4 As Long = 14396230 'RGB(70, 171, 219)
Public Const rampOcean5 As Long = 13800982 '?RGB(22, 150, 210) in Immediate window
Public Const rampOcean6 As Long = 10383634 'RGB(18, 113, 158)
Public Const rampOcean7 As Long = 6966282  'RGB(10, 76, 106)

Public Const giRGBgridlinesweb As Long = 14540253   'RGB(221, 221, 221)
Public Const giRGBgridlinesprint As Long = 14277081 'RGB(217, 217, 217)

'data series constants
Public Const seriesGapWidth As Double = 33
Public Const seriesOverlap As Double = 0

'chart area constants
Public Const titleFontSize As Double = 20
Public Const subTitleFontSize As Double = 16
Public Const axisFontSize As Double = 12

'=== modChartBuilder: OuterFormat ===
Public Const plotAreaHeight As Long = 310
Public Const plotAreaWidth As Long = 740
Public Const plotAreaTop_legend As Long = 85        ' single-series or multi-series with legend
Public Const plotAreaTop_noLegend As Long = 70      ' multi-series, no legend
Public Const plotAreaLeft As Long = 2

'=== modChartBuilder: FormatXAxisTitle ===
Public Const xAxisTitle_legendGap As Long = 15      ' spacing between title box and legend
Public Const xAxisTitle_singleOffset As Long = 3    ' upward nudge for single-series charts
Public Const legend_top As Long = 45
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

'=== modChartBuilder: FormatTitle ===
Public Const titleBoxWidth As Long = 250
Public Const titleBoxHeight As Long = 30
Public Const subtitleBoxTop As Long = 30
Public Const subtitleBoxHeight As Long = 25
Public Const titleBoxNudge As Long = 4              ' pixel nudge applied to top/left for alignment
Public Const yAxisLabel_legendTop As Long = 60
Public Const yAxisLabel_singleTop As Long = 52
Public Const yAxisLabel_multiTop As Long = 40
Public Const yAxisLabel_legendHeight As Long = 20
Public Const yAxisLabel_noLegendHeight As Long = 18

'=== modChartBuilder: FormatGridlines ===
Public Const gridlineWeight As Double = 1

'=== modChartBuilder: FormatXAxis ===
Public Const tickLabelSize_print As Double = 8.5
Public Const axisLineWeight As Double = 1

'=== modChartGeneric: SetPieChartSizeandTitle ===
Public Const pieTitleFontSize As Double = 18
Public Const pieTitleBoxHeight As Long = 25         ' pie title box height (smaller than standard)
Public Const pieSubtitleFontSize As Double = 14
Public Const pieSubtitleBoxTop As Long = 25
Public Const pieSubtitleBoxHeight As Long = 20
Public Const pieYAxisLabelBoxTop As Long = 45
Public Const pieTitleBoxHeight_print As Long = 15
Public Const piePlotAreaSize_legend As Long = 320   ' width and height (square) when legend present
Public Const piePlotAreaSize_noLegend As Long = 340 ' width and height (square) without legend
Public Const piePlotAreaLeft_web As Long = 100
Public Const piePlotAreaTop_web As Long = 40
Public Const piePlotAreaHeight_print As Long = 225
Public Const piePlotTopRatio_web As Double = 0.75   ' vertical centering ratio (web)
Public Const piePlotTopRatio_print As Double = 0.5  ' vertical centering ratio (print)
Public Const pieLegendTop_web As Long = 60
Public Const pieLegendTop_print As Long = 0

'=== modChartPie: BuildPieChart ===
Public Const pieErrorBoxWidth As Long = 500
Public Const pieErrorBoxHeight As Long = 40
Public Const pieErrorFontSize As Double = 10

'=== modChartGeneric: RemoveLegendandResize ===
Public Const removeLegend_webHeight As Long = 170
Public Const removeLegend_webTop As Long = 65
Public Const removeLegend_webWidth As Long = 300
Public Const removeLegend_webLeft As Long = 1
Public Const removeLegend_printHeight As Long = 225
Public Const removeLegend_printTop As Long = 20
Public Const removeLegend_printWidth As Long = 450
Public Const removeLegend_printLeft As Long = 5

'=== modChartGeneric: SlopeChartStyles ===
Public Const slopePlotWidth_web As Long = 700
Public Const slopePlotWidth_print As Long = 400
Public Const slopePlotLeft As Long = 40
Public Const slopePlotTop As Long = 80
Public Const slopeLegendTop_print As Long = 0
Public Const slopeLegend_printLeftPad As Long = 8

'=== modChartGeneric: DotPlotStyles and SlopeChartStyles (shared label sizes) ===
Public Const dataLabelFontSize_web As Double = 12   ' primary label size (web) — same as axisFontSize
Public Const dataLabelFontSize_secondary As Double = 11  ' secondary/normal label text (web)
Public Const dataLabelFontSize_print As Double = 9.5     ' primary label size (print)

'=== modExport ===
Public Const exportAppName As String = "INSO Chart Styles"
Public Const exportSection As String = "Chart Export"
Public Const exportSettingKey As String = "File Filter"
Public Const exportDefaultExt As String = "png"
Public Const exportDefaultName As String = "MyChart"
