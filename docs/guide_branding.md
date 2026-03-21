# Configuring the Add-in for a Company Brand

All brand-specific settings are isolated in three files: `modConfig.bas`, `modConfigColors.bas`, and `modEmbeddedImages.bas`. No other module needs to be touched for a standard white-label deployment.

---

## 1 — Organisation name: `modConfig.bas`

```vba
Public Const orgName As String = "COMPANY"
```

This single string appears in:
- The ribbon tab label in `CustomUI14.xml` (currently hardcoded as `"COMPANY Chart Styles"` — update that separately, see section 5)
- The Windows registry key used to persist the last-used export format: `"COMPANY Chart Styles"` → `"Chart Export"`

Change `"COMPANY"` to the client's short name (no spaces recommended — it becomes a registry key segment).

---

## 2 — Fonts: `modConfig.bas`

```vba
Public Const fontPrimary As String = "Calibri"
Public Const fontPrimaryItalic As String = "Calibri Italic"
```

Replace with the brand's chart font. `fontPrimary` is applied to all chart text. `fontPrimaryItalic` is used for the Y-axis label and the X-axis title placeholder.

> The font must be installed on every machine that uses the add-in. If the font is missing, Excel silently substitutes a fallback. If the brand font has no separate italic face, set both constants to the same font name.

---

## 3 — Chart canvas size: `modConfig.bas`

```vba
Public Const chartWidth As Double  = 600   ' points — 8.33" at 72dpi
Public Const chartHeight As Double = 600   ' points — 8.33" at 72dpi
```

1 point = 1/72 inch. Common target sizes:

| Use case | Width × Height (pt) |
|---|---|
| Web/screen, square | 600 × 600 |
| Web/screen, 4:3 | 640 × 480 |
| Web/screen, 16:9 | 640 × 360 |
| Print, half-page | 432 × 288 |

When you change canvas size, the plot area constants will also need proportional adjustment (see section 4).

---

## 4 — Layout constants: `modConfig.bas`

The layout constants describe where every element sits within the chart canvas, measured in points from the top-left corner. They are organised by the pipeline function that uses them.

```
Chart canvas (chartWidth × chartHeight)
│
├── TitleBox          top-left, width = titleBoxWidth
├── SubTitleBox       top = subtitleBoxTop
├── YAxisLabelBox     top = yAxisLabel_*Top
│
├── PlotArea          top = plotAreaTop_*, left = plotAreaLeft
│   └── XAxisTitle    placed below plot area inner boundary
│
├── Legend            top = legend_top, left = legend_leftPad (when present)
│
├── SourceBox         anchored to bottom-left
└── LogoImage         anchored to bottom-right
```

### Adjusting layout after a size change

The most common adjustment is a proportional scale. If you change `chartHeight` from 600 to 480 (80%), scale all vertical constants by 0.8 and test visually. The constants you are most likely to need to tune are:

| Constant | What it controls |
|---|---|
| `plotAreaHeight` | Height of the data area |
| `plotAreaTop_default` | Top of data area (with legend or single series) |
| `plotAreaTop_noLegend` | Top of data area (multi-series, no legend) |
| `legend_top` | Vertical position of the legend |
| `subtitleBoxTop` | Gap between title and subtitle |

The comment block at the top of the `OuterFormat` section in `modConfig.bas` explains the interdependencies.

---

## 5 — Ribbon tab label: `CustomUI14.xml`

The ribbon tab label is hardcoded in the XML and is not read from `orgName` at runtime (Excel does not support VBA expressions in ribbon XML):

```xml
<tab id="Tab1" label="COMPANY Chart Styles">
```

Change `"COMPANY Chart Styles"` to the client's tab label. Also update every `supertip` attribute, which currently reads `"Style a chart following the COMPANY standards"`.

A quick find-and-replace of `COMPANY` in the XML handles both.

---

## 6 — Brand colours: `modConfigColors.bas`

### Categorical colours

The seven named colours are used for multi-series charts:

```vba
Public Const colorData1 As Long = 12285696     'Ocean    RGB(0, 119, 187)
Public Const colorData2 As Long = 6719743      'Coral    RGB(255, 136, 102)
Public Const colorData3 As Long = 16764023     'Sky      RGB(119, 204, 255)
Public Const colorData4 As Long = 8952064      'Pine     RGB(0, 153, 136)
Public Const colorData5 As Long = 3399167      'Gold     RGB(255, 221, 51)
Public Const colorData6 As Long = 17578        'Rust     RGB(170, 68, 0)
Public Const colorData7 As Long = 15636906     'Lavender RGB(170, 153, 238
```

### Converting an RGB value to a VBA Long

Excel's object model stores colours as `Long` integers in BGR byte order (Blue, Green, Red — the reverse of standard RGB). The formula is:

```
Long = Blue × 65536 + Green × 256 + Red
```

You can also use Excel's built-in `RGB()` function in the Immediate Window to get the value:

```vba
' In the VBE Immediate Window (Ctrl+G):
?RGB(0, 119, 187)
' → prints the Long value to use as the constant
```

Paste that number as the constant value and add the human-readable RGB as a comment:

```vba
Public Const colorPrimary As Long = 12285696  'RGB(0, 119, 187)
```

> The comment is documentation only — it does not affect the value. Keep it accurate.

### Colour ramps

Each of the seven hues has a 7-step sequential ramp (`ramp<Hue>1` through `ramp<Hue>7`, where 1 = lightest and 7 = darkest). If the brand has fewer signature hues, replace the unused ramp sets with monochrome or neutral scales — the names must remain the same because `LoadPalette` in `modRamp.bas` references them.

---

## 7 — Logo: `modEmbeddedImages.bas`

The logo is embedded as a Base64-encoded string directly in the VBA module. This avoids any file-path dependency — the add-in is self-contained.

### Step 1 — Prepare the image

The logo should be:
- **PNG or SVG** — both are supported by `Shapes.AddPicture`. PNG is safer across Excel versions.
- **Transparent background** — the logo appears over the chart background, so transparency matters.
- **Square or near-square** — the sizing is controlled by `logoHeightScale` and `logoAspectRatio` in `modConfig.bas`. Adjust these if the logo is a wide horizontal lockup rather than a square mark.

### Step 2 — Convert to Base64 (PowerShell)

Run the following PowerShell script. It reads the image file, encodes it as Base64, splits it into 512-character chunks (to respect VBA's line-length limit), and writes the formatted VBA code to a text file.

```powershell
# Replace with your logo file path
$imagePath = "C:\path\to\your\logo.png"

# Read and encode
$bytes = [IO.File]::ReadAllBytes($imagePath)
$b64   = [Convert]::ToBase64String($bytes)

# Build the VBA function body
$sb = [System.Text.StringBuilder]::new()
$null = $sb.AppendLine("Public Function LogoPNG_Base64() As String")
$null = $sb.AppendLine("    ' Returns the embedded image as a Base64 string (split to avoid line-length limits)")
$null = $sb.AppendLine("    Dim s As String")

for ($i = 0; $i -lt $b64.Length; $i += 512) {
    $chunk = $b64.Substring($i, [Math]::Min(512, $b64.Length - $i))
    $null = $sb.AppendLine("  s = s & _")
    $null = $sb.AppendLine("`"$chunk`"")
}

$null = $sb.AppendLine("    LogoPNG_Base64 = s")
$null = $sb.AppendLine("End Function")

# Write to file
$outPath = [IO.Path]::ChangeExtension($imagePath, "txt")
[IO.File]::WriteAllText($outPath, $sb.ToString(), [Text.Encoding]::UTF8)
Write-Host "Written to $outPath"
```

This produces a `.txt` file containing the `LogoPNG_Base64` function body. Open it, copy the full content, and use it to replace the body of `LogoPNG_Base64` in `modEmbeddedImages.bas`.

### Step 3 — Update aspect ratio

In `modConfig.bas`, update `logoAspectRatio` to match the new logo's width-to-height ratio:

```vba
Public Const logoAspectRatio As Double = 1.8   ' logo width = aspectRatio × height
```

For a 300×100 pixel logo: `300 / 100 = 3.0`.
For a 200×200 square logo: `200 / 200 = 1.0`.

Also consider `logoHeightScale` — the logo height as a fraction of the chart height:

```vba
Public Const logoHeightScale As Double = 0.1   ' 10% of chartHeight
```

Larger logos (wide horizontal lockups) often need `0.07` to `0.09` to avoid dominating the chart. Square marks can stay at `0.1` or go slightly larger.

### Step 4 — Test the logo

Rebuild the `.xlam`, apply any chart, and inspect the bottom-right corner. If the logo is positioned incorrectly, adjust `logoMarginRight` and `logoMarginBottom` in `modConfig.bas`.

---

## 8 — Full constant reference

All values are in Excel points unless noted. These are the defaults at 600 × 600 pt canvas; scale proportionally when changing canvas size (see section 3).

### Fonts

| Constant | Description | Default |
|---|---|---|
| `fontPrimary` | All chart text | `"Calibri"` |
| `fontPrimaryItalic` | Y-axis label, x-axis title | `"Calibri Italic"` |
| `titleFontSize` | Chart title | `24` pt |
| `subTitleFontSize` | Subtitle | `20` pt |
| `axisFontSize` | Axis tick labels, legend | `16` pt |

### Series (bar and column)

| Constant | Description | Default |
|---|---|---|
| `seriesGapWidth` | Gap between bar/column groups | `33`% |
| `seriesOverlap` | Overlap between bars within a group | `0`% |

### Plot area

| Constant | Description | Default |
|---|---|---|
| `plotAreaHeight` | Plot area height | `408` pt |
| `plotAreaWidth` | Plot area width | `973` pt |
| `plotAreaTop_default` | Top — single-series or multi-series with legend | `158` pt |
| `plotAreaTop_noLegend` | Top — multi-series, no legend | `116` pt |
| `plotAreaLeft` | Left offset | `3` pt |
| `plotArea_noLegendSingleHeight` | Height — single series, no legend | `421` pt |
| `plotArea_noLegendSingleTop` | Top — single series, no legend | `105` pt |
| `plotArea_noLegendMultiHeight` | Height — multi-series, no legend | `460` pt |
| `plotArea_noLegendMultiTop` | Top — multi-series, no legend | `79` pt |

### Legend and x-axis title

| Constant | Description | Default |
|---|---|---|
| `xAxisTitle_plotGap` | Gap between plot area base and x-axis title | `20` pt |
| `legend_top` | Legend top offset | `92` pt |
| `legend_leftPad` | Legend left offset | `7` pt |

### Title boxes

| Constant | Description | Default |
|---|---|---|
| `titleBoxWidth` | Title and subtitle text box width | `394` pt |
| `titleBoxHeight` | Title text box height | `39` pt |
| `titleBoxNudge` | Top/left alignment nudge | `5` pt |
| `subtitleBoxTop` | Subtitle text box top offset | `53` pt |
| `subtitleBoxHeight` | Subtitle text box height | `33` pt |
| `yAxisLabel_legendTop` | Y-axis label top — legend present | `126` pt |
| `yAxisLabel_singleTop` | Y-axis label top — single series, no legend | `85` pt |
| `yAxisLabel_multiTop` | Y-axis label top — multi-series, no legend | `68` pt |
| `yAxisLabel_legendHeight` | Y-axis label height — legend present | `26` pt |
| `yAxisLabel_noLegendHeight` | Y-axis label height — no legend | `24` pt |

### Source / notes box

| Constant | Description | Default |
|---|---|---|
| `sourceBoxWidth` | Source/notes text box width | `230` pt |
| `sourceBoxHeight` | Source/notes text box height | `46` pt |
| `sourceTextFontSize` | Source/notes font size | `11` pt |
| `sourceBoxLeftNudge` | Left offset nudge | `5` pt |

### Logo

| Constant | Description | Default |
|---|---|---|
| `logoHeightScale` | Logo height as fraction of chart height | `0.1` (10%) |
| `logoAspectRatio` | Logo width = height × this value | `1.8` |
| `logoMarginRight` | Right margin | `10` pt |
| `logoMarginBottom` | Bottom margin | `8` pt |

### Gridlines and axes

| Constant | Description | Default |
|---|---|---|
| `gridlineWeight` | Major gridline stroke weight | `1` pt |
| `axisLineWeight` | X-axis line stroke weight | `1` pt |

### Pie and donut

| Constant | Description | Default |
|---|---|---|
| `piePlotAreaSize_legend` | Plot area size (square) — legend present | `421` pt |
| `piePlotAreaSize_noLegend` | Plot area size (square) — no legend | `447` pt |
| `piePlotAreaLeft_web` | Plot area left offset | `131` pt |
| `piePlotAreaTop_web` | Plot area initial top offset | `53` pt |
| `piePlotTopRatio_web` | Vertical centering ratio | `0.75` |
| `pieLegendTop_web` | Legend top offset | `79` pt |

### Lollipop

| Constant | Description | Default |
|---|---|---|
| `lollipopGapWidth` | Bar group gap width | `150`% |
| `lollipopStickWeight` | Error bar (stick) line weight | `1.5` pt |

### Chart tools

| Constant | Description | Default |
|---|---|---|
| `removeLegend_webHeight` | Chart height after Remove Legend | `224` pt |
| `removeLegend_webTop` | Plot area top after Remove Legend | `85` pt |
| `removeLegend_webWidth` | Plot area width after Remove Legend | `394` pt |
| `removeLegend_webLeft` | Plot area left after Remove Legend | `1` pt |
| `labelLastPointPlotWidthInset` | Plot area width reduction for end labels | `66` pt |
| `labelLastPointPlotTop` | Plot area top after Label Last Point | `105` pt |
| `labelLastPointPlotWidthRatio_web` | Plot area width ratio | `0.98` |
| `labelLastPointTitleNudge` | Title box nudge | `−13` pt |

### Export

| Constant | Description | Default |
|---|---|---|
| `exportDefaultExt` | Default file extension | `"png"` |
| `exportDefaultName` | Default file name | `"MyChart"` |

---

## 9 — Checklist for a complete white-label deployment

| Item | Location | Status |
|---|---|---|
| `orgName` | `modConfig.bas:14` | |
| Font name(s) | `modConfig.bas:9–10` | |
| Canvas size | `modConfig.bas:5–6` | |
| Layout constants (if canvas resized) | `modConfig.bas:31–117` | |
| Ribbon tab label and supertips | `CustomUI14.xml` | |
| Categorical colours (7) | `modConfigColors.bas:20–26` | |
| Sequential ramp colours (49) | `modConfigColors.bas:29–90` | |
| Logo Base64 | `modEmbeddedImages.bas` | |
| Logo aspect ratio | `modConfig.bas:48` | |
| Logo scale and margins | `modConfig.bas:47–50` | |
| Colour names (if renaming hues) | `modRamp.bas`, `modFormatFill.bas`, `CustomUI14.xml` | |
