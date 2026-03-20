Attribute VB_Name = "modConfigColors"
'==== Module: modConfigColors ====
'=== Brand colors ===
' Note: colorPrimaryBlue, colorDarkBlue, colorLightGrey are defined for completeness
' but are not currently referenced in code. Reserved for future ribbon buttons.
Public Const colorPrimaryBlue As Long = 10963739    'RGB(27, 75, 167)
Public Const colorDarkBlue As Long = 2888711        'RGB(7, 20, 44)
Public Const colorBlack As Long = 655874            'RGB(2, 2, 10)
Public Const colorLightGrey As Long = 16382457      'RGB(249, 249, 249)

'=== Neutral colors ===
' Note: colorSteel and colorAsh are defined for completeness but not currently referenced in code.
Public Const colorSilver As Long = 13421772     'RGB(204, 204, 204)
Public Const colorSteel As Long = 12303291      'RGB(187, 187, 187)
Public Const colorAsh As Long = 10263708        'RGB(156, 156, 156)
Public Const colorWhite As Long = 16777215      'RGB(255, 255, 255)

'=== Data colors ===
Public Const colorOcean As Long = 12285696      '?RGB(0, 119, 187)
Public Const colorCoral As Long = 6719743       'RGB(255, 136, 102)
Public Const colorSky As Long = 16764023        'RGB(119, 204, 255)
Public Const colorPine As Long = 8952064        'RGB(0, 153, 136)
Public Const colorGold As Long = 3399167        'RGB(255, 221, 51)
Public Const colorRust As Long = 17578          'RGB(170, 68, 0)
Public Const colorLavender As Long = 15636906   'RGB(170, 153, 238)

'== Color ramp ==
' Ocean ramp (sequential palette for single-hue charts)
Public Const rampOcean1 As Long = 15984847 'RGB(207, 232, 243)
Public Const rampOcean2 As Long = 15520930 'RGB(162, 212, 236)
Public Const rampOcean3 As Long = 14860147 'RGB(115, 191, 226)
Public Const rampOcean4 As Long = 14396230 'RGB(70, 171, 219)
Public Const rampOcean5 As Long = 13800982 '?RGB(22, 150, 210) in Immediate window
Public Const rampOcean6 As Long = 10383634 'RGB(18, 113, 158)
Public Const rampOcean7 As Long = 6966282  'RGB(10, 76, 106)

' Coral ramp
Public Const rampCoral1 As Long = 15791103 'RGB(255, 243, 240)
Public Const rampCoral2 As Long = 14739455 'RGB(255, 231, 224)
Public Const rampCoral3 As Long = 12767231 'RGB(255, 207, 194)
Public Const rampCoral4 As Long = 10729727 'RGB(255, 184, 163)
Public Const rampCoral5 As Long = 6719743  'RGB(255, 136, 102)
Public Const rampCoral6 As Long = 5402060  'RGB(204, 109, 82)
Public Const rampCoral7 As Long = 4018841  'RGB(153, 82, 61)

' Sky ramp
Public Const rampSky1 As Long = 16774628   'RGB(228, 245, 255)
Public Const rampSky2 As Long = 16772041   'RGB(201, 235, 255)
Public Const rampSky3 As Long = 16769197   'RGB(173, 224, 255)
Public Const rampSky4 As Long = 16766610   'RGB(146, 214, 255)
Public Const rampSky5 As Long = 16764023   'RGB(119, 204, 255)
Public Const rampSky6 As Long = 13411167   'RGB(95, 163, 204)
Public Const rampSky7 As Long = 10058311   'RGB(71, 122, 153)

' Pine ramp
Public Const rampPine1 As Long = 15988198  'RGB(230, 245, 243)
Public Const rampPine2 As Long = 15199180  'RGB(204, 235, 231)
Public Const rampPine3 As Long = 13620889  'RGB(153, 214, 207)
Public Const rampPine4 As Long = 12108390  'RGB(102, 194, 184)
Public Const rampPine5 As Long = 8952064   'RGB(0, 153, 136)
Public Const rampPine6 As Long = 5397504   'RGB(0, 92, 82)
Public Const rampPine7 As Long = 3554560   'RGB(0, 61, 54)

' Gold ramp
Public Const rampGold1 As Long = 15465727  'RGB(255, 252, 235)
Public Const rampGold2 As Long = 14088447  'RGB(255, 248, 214)
Public Const rampGold3 As Long = 11399679  'RGB(255, 241, 173)
Public Const rampGold4 As Long = 8776703   'RGB(255, 235, 133)
Public Const rampGold5 As Long = 3399167   'RGB(255, 221, 51)
Public Const rampGold6 As Long = 48096     'RGB(224, 187, 0)
Public Const rampGold7 As Long = 34979     'RGB(163, 136, 0)

' Rust ramp
Public Const rampRust1 As Long = 13425390  'RGB(238, 218, 204)
Public Const rampRust2 As Long = 10073309  'RGB(221, 180, 153)
Public Const rampRust3 As Long = 6721484   'RGB(204, 143, 102)
Public Const rampRust4 As Long = 3369403   'RGB(187, 105, 51)
Public Const rampRust5 As Long = 17578     'RGB(170, 68, 0)
Public Const rampRust6 As Long = 10598     'RGB(102, 41, 0)
Public Const rampRust7 As Long = 6980      'RGB(68, 27, 0)

' Lavender ramp
Public Const rampLavender1 As Long = 16643575 'RGB(247, 245, 253)
Public Const rampLavender2 As Long = 16575470 'RGB(238, 235, 252)
Public Const rampLavender3 As Long = 16307933 'RGB(221, 214, 248)
Public Const rampLavender4 As Long = 16106188 'RGB(204, 194, 245)
Public Const rampLavender5 As Long = 15636906 'RGB(170, 153, 238)
Public Const rampLavender6 As Long = 12483208 'RGB(136, 122, 190)
Public Const rampLavender7 As Long = 9395302  'RGB(102, 92, 143)