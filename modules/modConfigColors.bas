Attribute VB_Name = "modConfigColors"
Option Explicit
'==== Module: modConfigColors ====
'=== Brand colors ===
' Note: colorBrand1, colorBrand2, colorBrand4 are defined for completeness
' but are not currently referenced in code. Reserved for future ribbon buttons.
Public Const colorBrand1 As Long = 10963739    'Primary Blue  RGB(27, 75, 167)
Public Const colorBrand2 As Long = 2888711     'Dark Blue     RGB(7, 20, 44)
Public Const colorBrand3 As Long = 655874      'Black        RGB(2, 2, 10)
Public Const colorBrand4 As Long = 16382457    'Light Grey    RGB(249, 249, 249)

'=== Neutral colors ===
' Note: colorNeutral2 and colorNeutral3 are defined for completeness but not currently referenced in code.
Public Const colorNeutral1 As Long = 13421772  'Silver  RGB(204, 204, 204)
Public Const colorNeutral2 As Long = 12303291  'Steel   RGB(187, 187, 187)
Public Const colorNeutral3 As Long = 10263708  'Ash     RGB(156, 156, 156)
Public Const colorNeutral4 As Long = 16777215  'White   RGB(255, 255, 255)

'=== Data colors ===
Public Const colorData1 As Long = 12285696     'Ocean    RGB(0, 119, 187)
Public Const colorData2 As Long = 6719743      'Coral    RGB(255, 136, 102)
Public Const colorData3 As Long = 16764023     'Sky      RGB(119, 204, 255)
Public Const colorData4 As Long = 8952064      'Pine     RGB(0, 153, 136)
Public Const colorData5 As Long = 3399167      'Gold     RGB(255, 221, 51)
Public Const colorData6 As Long = 17578        'Rust     RGB(170, 68, 0)
Public Const colorData7 As Long = 15636906     'Lavender RGB(170, 153, 238)

'== Color ramp ==
' rampA = Ocean (sequential palette for single-hue charts)
Public Const rampA1 As Long = 15984847 'RGB(207, 232, 243)
Public Const rampA2 As Long = 15520930 'RGB(162, 212, 236)
Public Const rampA3 As Long = 14860147 'RGB(115, 191, 226)
Public Const rampA4 As Long = 14396230 'RGB(70, 171, 219)
Public Const rampA5 As Long = 13800982 'RGB(22, 150, 210)
Public Const rampA6 As Long = 10383634 'RGB(18, 113, 158)
Public Const rampA7 As Long = 6966282  'RGB(10, 76, 106)

' rampB = Coral
Public Const rampB1 As Long = 15791103 'RGB(255, 243, 240)
Public Const rampB2 As Long = 14739455 'RGB(255, 231, 224)
Public Const rampB3 As Long = 12767231 'RGB(255, 207, 194)
Public Const rampB4 As Long = 10729727 'RGB(255, 184, 163)
Public Const rampB5 As Long = 6719743  'RGB(255, 136, 102)
Public Const rampB6 As Long = 5402060  'RGB(204, 109, 82)
Public Const rampB7 As Long = 4018841  'RGB(153, 82, 61)

' rampC = Sky
Public Const rampC1 As Long = 16774628   'RGB(228, 245, 255)
Public Const rampC2 As Long = 16772041   'RGB(201, 235, 255)
Public Const rampC3 As Long = 16769197   'RGB(173, 224, 255)
Public Const rampC4 As Long = 16766610   'RGB(146, 214, 255)
Public Const rampC5 As Long = 16764023   'RGB(119, 204, 255)
Public Const rampC6 As Long = 13411167   'RGB(95, 163, 204)
Public Const rampC7 As Long = 10058311   'RGB(71, 122, 153)

' rampD = Pine
Public Const rampD1 As Long = 15988198  'RGB(230, 245, 243)
Public Const rampD2 As Long = 15199180  'RGB(204, 235, 231)
Public Const rampD3 As Long = 13620889  'RGB(153, 214, 207)
Public Const rampD4 As Long = 12108390  'RGB(102, 194, 184)
Public Const rampD5 As Long = 8952064   'RGB(0, 153, 136)
Public Const rampD6 As Long = 5397504   'RGB(0, 92, 82)
Public Const rampD7 As Long = 3554560   'RGB(0, 61, 54)

' rampE = Gold
Public Const rampE1 As Long = 15465727  'RGB(255, 252, 235)
Public Const rampE2 As Long = 14088447  'RGB(255, 248, 214)
Public Const rampE3 As Long = 11399679  'RGB(255, 241, 173)
Public Const rampE4 As Long = 8776703   'RGB(255, 235, 133)
Public Const rampE5 As Long = 3399167   'RGB(255, 221, 51)
Public Const rampE6 As Long = 48096     'RGB(224, 187, 0)
Public Const rampE7 As Long = 34979     'RGB(163, 136, 0)

' rampF = Rust
Public Const rampF1 As Long = 13425390  'RGB(238, 218, 204)
Public Const rampF2 As Long = 10073309  'RGB(221, 180, 153)
Public Const rampF3 As Long = 6721484   'RGB(204, 143, 102)
Public Const rampF4 As Long = 3369403   'RGB(187, 105, 51)
Public Const rampF5 As Long = 17578     'RGB(170, 68, 0)
Public Const rampF6 As Long = 10598     'RGB(102, 41, 0)
Public Const rampF7 As Long = 6980      'RGB(68, 27, 0)

' rampG = Lavender
Public Const rampG1 As Long = 16643575 'RGB(247, 245, 253)
Public Const rampG2 As Long = 16575470 'RGB(238, 235, 252)
Public Const rampG3 As Long = 16307933 'RGB(221, 214, 248)
Public Const rampG4 As Long = 16106188 'RGB(204, 194, 245)
Public Const rampG5 As Long = 15636906 'RGB(170, 153, 238)
Public Const rampG6 As Long = 12483208 'RGB(136, 122, 190)
Public Const rampG7 As Long = 9395302  'RGB(102, 92, 143)
