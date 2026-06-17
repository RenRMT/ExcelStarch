Attribute VB_Name = "modConfigColors"
Option Explicit
'==== Module: modConfigColors ====
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
' as near-white. Used by modRamp.BuildDivergingRamp.
Public Const colorBrand4 As Long = 16382457    'Neutral centre #F9F9F9 RGB(249, 249, 249)

'=== Neutral colors ===
' Note: colorNeutral3 is defined for completeness but not currently referenced in code.
Public Const colorNeutral1 As Long = 15527148  'Platinum RGB(236, 236, 236) #ECECEC
Public Const colorNeutral2 As Long = 12105912  'Steel    RGB(184, 184, 184) #B8B8B8
Public Const colorNeutral3 As Long = 10263708  'Ash      RGB(156, 156, 156) - reserved
Public Const colorNeutral4 As Long = 16777215  'White    RGB(255, 255, 255) - used by name-lookup palette (modFormatFill)
Public Const colorWhite As Long = colorNeutral4  'Semantic alias - use this for axis/border white styling

'=== Data colors (contrasting order) ===
Public Const colorData1 As Long = 7833651      'Teal    #338877 RGB(51, 136, 119)
Public Const colorData2 As Long = 7855615      'Jasmine #FFDD77 RGB(255, 221, 119)
Public Const colorData3 As Long = 7811874      'Navy    #223377 RGB(34, 51, 119)
Public Const colorData4 As Long = 7968767      'Coral   #FF9779 RGB(255, 151, 121)
Public Const colorData5 As Long = 16764040     'Sky     #88CCFF RGB(136, 204, 255)
Public Const colorData6 As Long = 4471731      'Cherry  #B33B44 RGB(179, 59, 68)
Public Const colorData7 As Long = 14531583     'Blush   #FFBBDD RGB(255, 187, 221)
Public Const colorData8 As Long = 11154312     'Indigo  #8833AA RGB(136, 51, 170)

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

' rampC = Navy
Public Const rampC1 As Long = 15854569  '#e9ebf1 RGB(233, 235, 241)
Public Const rampC2 As Long = 14997203  '#d3d6e4 RGB(211, 214, 228)
Public Const rampC3 As Long = 13217191  '#a7adc9 RGB(167, 173, 201)
Public Const rampC4 As Long = 11371898  '#7a85ad RGB(122, 133, 173)
Public Const rampC5 As Long = 9591886   '#4e5c92 RGB(78, 92, 146)
Public Const rampC6 As Long = 7811874   '#223377 RGB(34, 51, 119)
Public Const rampC7 As Long = 6236443   '#1b295f RGB(27, 41, 95)
Public Const rampC8 As Long = 4661012   '#141f47 RGB(20, 31, 71)
Public Const rampC9 As Long = 3150862   '#0e1430 RGB(14, 20, 48)
Public Const rampC10 As Long = 1575431  '#070a18 RGB(7, 10, 24)

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
Public Const rampF1 As Long = 15526903  '#f7ebec RGB(247, 235, 236)
Public Const rampF2 As Long = 14342384  '#f0d8da RGB(240, 216, 218)
Public Const rampF3 As Long = 11842017  '#e1b1b4 RGB(225, 177, 180)
Public Const rampF4 As Long = 9406929   '#d1898f RGB(209, 137, 143)
Public Const rampF5 As Long = 6906562   '#c26269 RGB(194, 98, 105)
Public Const rampF6 As Long = 4471731   '#B33B44 RGB(179, 59, 68)
Public Const rampF7 As Long = 3551119   '#8f2f36 RGB(143, 47, 54)
Public Const rampF8 As Long = 2696043   '#6b2329 RGB(107, 35, 41)
Public Const rampF9 As Long = 1775688   '#48181b RGB(72, 24, 27)
Public Const rampF10 As Long = 920612   '#240c0e RGB(36, 12, 14)

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

' rampH = Indigo
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
