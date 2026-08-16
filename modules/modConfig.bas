Attribute VB_Name = "modConfig"
Option Explicit
'==== Module: modConfig ====
' Brand and user-editable settings - the values you adapt to your own house style.
' Everything here can be safely changed: brand/data colours and ramps, chart canvas
' dimensions, font types & sizes, font colours, logo size, margins, series spacing,
' and the default title/subtitle placeholder texts. The chart layout is generally
' responsive to changes you make here (you might need to test font sizes a little).
'
' Computed/derived values and the chart-engine defaults live in modConfigDerived;
' you should not need to touch that module unless you are changing the chart engine
' itself.


' +---------------------------------------------------------+
' |  BRAND COLOURS                                          |
' |  Edit these to match your house palette.                |
' +---------------------------------------------------------+
'=== Brand colors ===
' Note: colorBrand1, colorBrand2 are defined for completeness
' but are not currently referenced in code. Reserved for future ribbon buttons.
Public Const colorBrand1 As Long = 5793564     'Primary Teal #1C6758  RGB(28, 103, 88)
Public Const colorBrand2 As Long = 5394995     'Dark Teal    #335252  RGB(51, 82, 82)
Public Const colorBrand3 As Long = 1644825     'Black        #191919  RGB(25, 25, 25)
' Brand Light Grey (#F3F3F3) - reserved for future ribbon buttons.
Public Const colorBrandLightGrey As Long = 15987699 'Light Grey #F3F3F3 RGB(243, 243, 243)
' colorBrand4 is the diverging-ramp neutral centre (#F9F9F9). This is intentionally
' a touch lighter than the brand Light Grey (#F3F3F3 above) so the centre series reads
' as near-white. Used by modColorRamp.BuildDivergingRamp.
Public Const colorBrand4 As Long = 16382457    'Neutral centre #F9F9F9 RGB(249, 249, 249)

'=== Neutral colors ===
' Note: colorNeutral3 is defined for completeness but not currently referenced in code.
Public Const colorNeutral1 As Long = 15527148  'Platinum RGB(236, 236, 236) #ECECEC
Public Const colorNeutral2 As Long = 12105912  'Steel    RGB(184, 184, 184) #B8B8B8
Public Const colorNeutral3 As Long = 10263708  'Ash      RGB(156, 156, 156) - reserved
Public Const colorNeutral4 As Long = 16777215  'White    RGB(255, 255, 255) - used by name-lookup palette (modColorFill)
Public Const colorWhite As Long = colorNeutral4  'Semantic alias - use this for axis/border white styling

'=== Data colors (contrasting order) ===
Public Const colorData1 As Long = 7833651      'Teal    #338877 RGB(51, 136, 119)
Public Const colorData2 As Long = 7855615      'Jasmine #FFDD77 RGB(255, 221, 119)
Public Const colorData3 As Long = 10181419     'Baltic  #2B5B9B RGB(43, 91, 155)
Public Const colorData4 As Long = 7968767      'Coral   #FF9779 RGB(255, 151, 121)
Public Const colorData5 As Long = 16764040     'Sky     #88CCFF RGB(136, 204, 255)
Public Const colorData6 As Long = 4996284      'Cherry  #BC3C4C RGB(188, 60, 76)
Public Const colorData7 As Long = 14531583     'Blush   #FFBBDD RGB(255, 187, 221)
Public Const colorData8 As Long = 11154312     'Violet  #8833AA RGB(136, 51, 170)

'== Color ramp ==
' Each ramp is a 10-step sequential palette (1 = lightest .. 10 = darkest).
' rampA = Teal
Public Const rampA1 As Long = 15856619  '#ebf3f1 RGB(235, 243, 241)
Public Const rampA2 As Long = 15001558  '#d6e7e4 RGB(214, 231, 228)
Public Const rampA3 As Long = 13225901  '#adcfc9 RGB(173, 207, 201)
Public Const rampA4 As Long = 11384965  '#85b8ad RGB(133, 184, 173)
Public Const rampA5 As Long = 9609308   '#5ca092 RGB(92, 160, 146)
Public Const rampA6 As Long = 7833651   '#338877 RGB(51, 136, 119)
Public Const rampA7 As Long = 6253865   '#296d5f RGB(41, 109, 95)
Public Const rampA8 As Long = 4674079   '#1f5247 RGB(31, 82, 71)
Public Const rampA9 As Long = 3159572   '#143630 RGB(20, 54, 48)
Public Const rampA10 As Long = 1579786  '#0a1b18 RGB(10, 27, 24)

' rampB = Jasmine
Public Const rampB1 As Long = 15858943  '#fffcf1 RGB(255, 252, 241)
Public Const rampB2 As Long = 15005951  '#fff8e4 RGB(255, 248, 228)
Public Const rampB3 As Long = 13234687  '#fff1c9 RGB(255, 241, 201)
Public Const rampB4 As Long = 11398143  '#ffebad RGB(255, 235, 173)
Public Const rampB5 As Long = 9626879   '#ffe492 RGB(255, 228, 146)
Public Const rampB6 As Long = 7855615   '#FFDD77 RGB(255, 221, 119)
Public Const rampB7 As Long = 6271436   '#ccb15f RGB(204, 177, 95)
Public Const rampB8 As Long = 4687257   '#998547 RGB(153, 133, 71)
Public Const rampB9 As Long = 3168358   '#665830 RGB(102, 88, 48)
Public Const rampB10 As Long = 1584179  '#332c18 RGB(51, 44, 24)

' rampC = Baltic
Public Const rampC1 As Long = 16117738  '#eaeff5 RGB(234, 239, 245)
Public Const rampC2 As Long = 15458005  '#d5deeb RGB(213, 222, 235)
Public Const rampC3 As Long = 14138794  '#aabdd7 RGB(170, 189, 215)
Public Const rampC4 As Long = 12819840  '#809dc3 RGB(128, 157, 195)
Public Const rampC5 As Long = 11500629  '#557caf RGB(85, 124, 175)
Public Const rampC6 As Long = 10181419  '#2B5B9B RGB(43, 91, 155)
Public Const rampC7 As Long = 8145186   '#22497c RGB(34, 73, 124)
Public Const rampC8 As Long = 6108954   '#1a375d RGB(26, 55, 93)
Public Const rampC9 As Long = 4072465   '#11243e RGB(17, 36, 62)
Public Const rampC10 As Long = 2036233  '#09121f RGB(9, 18, 31)

' rampD = Coral
Public Const rampD1 As Long = 15922687  '#fff5f2 RGB(255, 245, 242)
Public Const rampD2 As Long = 15002367  '#ffeae4 RGB(255, 234, 228)
Public Const rampD3 As Long = 13227519  '#ffd5c9 RGB(255, 213, 201)
Public Const rampD4 As Long = 11518463  '#ffc1af RGB(255, 193, 175)
Public Const rampD5 As Long = 9743615   '#ffac94 RGB(255, 172, 148)
Public Const rampD6 As Long = 7968767   '#FF9779 RGB(255, 151, 121)
Public Const rampD7 As Long = 6388172   '#cc7961 RGB(204, 121, 97)
Public Const rampD8 As Long = 4807577   '#995b49 RGB(153, 91, 73)
Public Const rampD9 As Long = 3161190   '#663c30 RGB(102, 60, 48)
Public Const rampD10 As Long = 1580595  '#331e18 RGB(51, 30, 24)

' rampE = Sky
Public Const rampE1 As Long = 16775923  '#f3faff RGB(243, 250, 255)
Public Const rampE2 As Long = 16774631  '#e7f5ff RGB(231, 245, 255)
Public Const rampE3 As Long = 16772047  '#cfebff RGB(207, 235, 255)
Public Const rampE4 As Long = 16769208  '#b8e0ff RGB(184, 224, 255)
Public Const rampE5 As Long = 16766624  '#a0d6ff RGB(160, 214, 255)
Public Const rampE6 As Long = 16764040  '#88CCFF RGB(136, 204, 255)
Public Const rampE7 As Long = 13411181  '#6da3cc RGB(109, 163, 204)
Public Const rampE8 As Long = 10058322  '#527a99 RGB(82, 122, 153)
Public Const rampE9 As Long = 6705718   '#365266 RGB(54, 82, 102)
Public Const rampE10 As Long = 3352859  '#1b2933 RGB(27, 41, 51)

' rampF = Cherry
Public Const rampF1 As Long = 15592696  '#f8eced RGB(248, 236, 237)
Public Const rampF2 As Long = 14407922  '#f2d8db RGB(242, 216, 219)
Public Const rampF3 As Long = 12038628  '#e4b1b7 RGB(228, 177, 183)
Public Const rampF4 As Long = 9734871   '#d78a94 RGB(215, 138, 148)
Public Const rampF5 As Long = 7365577   '#c96370 RGB(201, 99, 112)
Public Const rampF6 As Long = 4996284   '#BC3C4C RGB(188, 60, 76)
Public Const rampF7 As Long = 4010134   '#96303d RGB(150, 48, 61)
Public Const rampF8 As Long = 3023985   '#71242e RGB(113, 36, 46)
Public Const rampF9 As Long = 1972299   '#4b181e RGB(75, 24, 30)
Public Const rampF10 As Long = 986150   '#260c0f RGB(38, 12, 15)

' rampG = Blush
Public Const rampG1 As Long = 16578815  '#fff8fc RGB(255, 248, 252)
Public Const rampG2 As Long = 16314879  '#fff1f8 RGB(255, 241, 248)
Public Const rampG3 As Long = 15852799  '#ffe4f1 RGB(255, 228, 241)
Public Const rampG4 As Long = 15455999  '#ffd6eb RGB(255, 214, 235)
Public Const rampG5 As Long = 14993919  '#ffc9e4 RGB(255, 201, 228)
Public Const rampG6 As Long = 14531583  '#FFBBDD RGB(255, 187, 221)
Public Const rampG7 As Long = 11638476  '#cc96b1 RGB(204, 150, 177)
Public Const rampG8 As Long = 8745113   '#997085 RGB(153, 112, 133)
Public Const rampG9 As Long = 5786470   '#664b58 RGB(102, 75, 88)
Public Const rampG10 As Long = 2893107  '#33252c RGB(51, 37, 44)

' rampH = Violet
Public Const rampH1 As Long = 16247795  '#f3ebf7 RGB(243, 235, 247)
Public Const rampH2 As Long = 15652583  '#e7d6ee RGB(231, 214, 238)
Public Const rampH3 As Long = 14527951  '#cfaddd RGB(207, 173, 221)
Public Const rampH4 As Long = 13403576  '#b885cc RGB(184, 133, 204)
Public Const rampH5 As Long = 12278944  '#a05cbb RGB(160, 92, 187)
Public Const rampH6 As Long = 11154312  '#8833AA RGB(136, 51, 170)
Public Const rampH7 As Long = 8923501   '#6d2988 RGB(109, 41, 136)
Public Const rampH8 As Long = 6692690   '#521f66 RGB(82, 31, 102)
Public Const rampH9 As Long = 4461622   '#361444 RGB(54, 20, 68)
Public Const rampH10 As Long = 2230811  '#1b0a22 RGB(27, 10, 34)


' +---------------------------------------------------------+
' |  USER SETTINGS                                          |
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
Public Const axisFontSize As Double = 18
Public Const sourceTextFontSize As Double = 14
Public Const generalFontSize As Double = 18

'Font colors
' colors as defined in the BRAND COLOURS section above.
' - generalFontColor: is currently not used for anything but kept for compatibility
Public Const titleFontColor As Long = colorBrand1
Public Const subTitleFontColor As Long = colorBrand2

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
Public Const donutHoleSize_percent As Long = 50  ' donut hole size %; valid 10-90. Excel's AddChart2 default is 75.
' Note: pieLegendTop is a derived value in modConfigDerived, since it derives
' from subtitleBoxTop/subtitleBoxHeight (VBA Const cannot forward-reference).

' Weights
'   - gridLineWeight:  weight of chart gridlines expressed in Excel points
'   - axisLineWeight:weight of chart axis lines expressed in Excel points
Public Const gridlineWeight As Double = 1
Public Const axisLineWeight As Double = 1

'=== Chart Actions settings ===
'Annotation box
Public Const annotationDefaultText As String = "Annotation"
Public Const annotationBoxWidth As Double = 120
Public Const annotationBoxHeight As Double = 30
Public Const annotationOffsetX As Double = 8     ' box left = point + offset
Public Const annotationOffsetY As Double = -8    ' box top  = point - offset (sit above)
Public Const annotationFontSize As Double = axisFontSize   ' reuse existing axis size
Public Const annotationFontColor As Long = axisFontColor   ' reuse existing axis colour


' === Box sizes ===
' Box sizes expressed as proportion of chart width / height
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
