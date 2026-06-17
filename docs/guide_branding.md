# Configuring the Add-in for a Company Brand

All brand-specific settings are isolated in three files: `modConfig.bas`, `modConfigColors.bas`, and `modEmbeddedImages.bas`. No other module needs to be touched for a standard white-label deployment.

---

## 1 — Organisation name: `modConfig.bas`

```vba
Public Const orgName As String = "COMPANY"
```

This string is used to derive the Windows-registry key that persists the last-used export format (`exportAppName = orgName & " Chart Styles"`). It does **not** drive the ribbon tab label — that is hardcoded separately in the XML (see section 5).

Change `"COMPANY"` to the client's short name. Avoid spaces, since it becomes part of a registry key.

---

## 2 — Fonts: `modConfig.bas`

```vba
Public Const fontPrimary As String = "Calibri"
Public Const fontPrimaryItalic As String = "Calibri Italic"
```

Replace with the brand's chart font. `fontPrimary` is applied to all chart text. `fontPrimaryItalic` is used for the y-axis label and x-axis title placeholder — set it to the same value as `fontPrimary` if you don't want italics. If the font is missing on the machine, Excel silently substitutes a fallback.

Font sizes are also in `modConfig.bas` (`titleFontSize`, `subTitleFontSize`, `figureFontSize`, `axisFontSize`, `sourceTextFontSize`). The placeholder text strings (`titleDefaultText`, `subtitleDefaultText`, etc.) reference these sizes, so update both together if you change them.

---

## 3 — Chart canvas size: `modConfig.bas`

```vba
Public Const chartWidth As Double  = 600   ' points (1pt = 1/72")
Public Const chartHeight As Double = 600   ' points
```

The layout is **responsive**: almost every other position and size is expressed as a proportion of `chartWidth`/`chartHeight` (see section 4), so changing the canvas dimensions rescales the whole chart automatically. You may still need to nudge a few font sizes by eye.

---

## 4 — How the layout system works: `modConfig.bas`

`modConfig.bas` is organised into two sections, and understanding the split is the key to safe customisation:

**Section 1 — User settings.** Values you are meant to edit: canvas size, fonts and font sizes, font colours, placeholder text, series gaps, logo settings, and the per-element **proportions**. Layout proportions are expressed as a fraction of a chart dimension — e.g. `titleBoxHeightProportion = 0.07` means the title box is 7% of `chartHeight`; `logoHeightScale = 0.1` makes the logo 10% of `chartHeight`. Because everything is proportional, the layout adapts when you change the canvas size.

**Section 2 — Derived constants.** Computed *from* Section 1 (e.g. `titleBoxHeight = chartHeight * titleBoxHeightProportion`, `logoTop = chartHeight - logoHeight - logoMarginBottom`). These exist so every absolute position stays consistent when a Section 1 value changes. **Do not edit these directly** — change the Section 1 input they derive from instead.

> **Source of truth:** rather than duplicate every constant here (such a table drifts out of date quickly), refer to `modConfig.bas` directly. It is heavily commented, and the Section 1 / Section 2 split tells you at a glance what is safe to edit. Tune Section 1 proportions and re-test visually; leave Section 2 alone.

A typical rebrand only touches a handful of Section 1 values: `orgName`, the fonts and sizes, the colours (section 6), and the logo (section 7). The proportions rarely need changing unless you are deviating from the default layout.

---

## 5 — Ribbon tab label and supertips: `CustomUI14.xml`

The ribbon tab label is hardcoded in the XML and is not read from `orgName` at runtime (Excel does not support VBA expressions in ribbon XML):

```xml
<tab id="Tab1" label="COMPANY Chart Styles">
```

Change `"COMPANY Chart Styles"` to the client's tab label. Also update every `supertip` attribute that reads `"Style a chart following the COMPANY standards"`.

A find-and-replace of `COMPANY` across the XML handles both.

---

## 6 — Brand colours: `modConfigColors.bas`

This module holds three colour families, all stored as VBA `Long` values:

- **Brand colours** (`colorBrand1`–`colorBrand4`, plus `colorBrandLightGrey`) — used for chart text (title, subtitle, figure/axis labels) and the diverging-ramp neutral centre. The defaults are a teal/black/grey set.
- **Neutral colours** (`colorNeutral1`–`colorNeutral4`, plus the `colorWhite` alias) — platinum/steel/ash/white, used for fallback fills and axis/border styling.
- **Data colours** (`colorData1`–`colorData8`) — the eight categorical hues for multi-series charts (Teal, Jasmine, Navy, Coral, Sky, Cherry, Blush, Indigo).

```vba
' Data colours (example — see modConfigColors.bas for the full set)
Public Const colorData1 As Long = 7833651      'Teal     RGB(51, 136, 119)
Public Const colorData2 As Long = 7855615      'Jasmine  RGB(255, 221, 119)
Public Const colorData3 As Long = 7811874      'Navy     RGB(34, 51, 119)
```

> Some brand/neutral constants (`colorBrand1`, `colorBrand2`, `colorBrandLightGrey`, `colorNeutral3`) are defined for completeness but not all are referenced in code — see the comments in the module. Note `colorBrand4` (#F9F9F9) is the diverging-ramp neutral centre and is intentionally lighter than the brand Light Grey (`colorBrandLightGrey`, #F3F3F3). **Do not rename any constant**; they are referenced by name across other modules. Change the *values* only.

### Converting an RGB value to a VBA Long

Excel stores colours as `Long` integers in BGR byte order (Blue, Green, Red — reverse of standard RGB):

```
Long = Blue × 65536 + Green × 256 + Red
```

The easiest way is Excel's built-in `RGB()` function in the Immediate Window (Ctrl+G):

```vba
?RGB(0, 119, 187)
' → prints the Long value to paste as the constant
```

Paste that number as the value and keep the human-readable RGB as a comment:

```vba
Public Const colorData1 As Long = 7833651  'Teal RGB(51, 136, 119)
```

The comment is documentation only and does not affect the value.

### Colour ramps

Each of the eight hues has a 10-step sequential ramp (`rampA1`–`rampA10` through `rampH1`–`rampH10`, where 1 = lightest and 10 = darkest). If the brand has fewer signature hues, replace the unused ramp sets with monochrome or neutral scales. As with all colour constants, **change values but not names** — they are referenced directly by `modRamp.bas`.

---

## 7 — Logo: `modEmbeddedImages.bas`

The logo is embedded as a Base64-encoded string directly in the VBA module to avoid dependencies on external files.

### Step 1 — Prepare the image

The logo should be:
- **PNG or SVG** — both are supported. Set `logoFileType` in `modConfig.bas` to match (`"svg"` by default). PNG is safest across Excel versions.
- **Transparent background** — the logo sits over the chart background.
- **Square or near-square** — sizing is controlled by `logoHeightScale` and `logoAspectRatio` in `modConfig.bas`.

### Step 2 — Convert to Base64 (PowerShell)

This script reads the image, encodes it as Base64, splits it into 512-character chunks (to respect VBA's line-length limit), and writes formatted VBA code to a text file:

```powershell
# Replace with your logo file path
$imagePath = "C:\path\to\your\logo.png"

$bytes = [IO.File]::ReadAllBytes($imagePath)
$b64   = [Convert]::ToBase64String($bytes)

$sb = [System.Text.StringBuilder]::new()
$null = $sb.AppendLine("Public Function LogoPNG_Base64() As String")
$null = $sb.AppendLine("    Dim s As String")

for ($i = 0; $i -lt $b64.Length; $i += 512) {
    $chunk = $b64.Substring($i, [Math]::Min(512, $b64.Length - $i))
    $null = $sb.AppendLine("  s = s & _")
    $null = $sb.AppendLine("`"$chunk`"")
}

$null = $sb.AppendLine("    LogoPNG_Base64 = s")
$null = $sb.AppendLine("End Function")

$outPath = [IO.Path]::ChangeExtension($imagePath, "txt")
[IO.File]::WriteAllText($outPath, $sb.ToString(), [Text.Encoding]::UTF8)
Write-Host "Written to $outPath"
```

Open the resulting `.txt`, copy its full contents, and replace the body of `LogoPNG_Base64` in `modEmbeddedImages.bas`.

### Step 3 — Update aspect ratio

In `modConfig.bas`, set `logoAspectRatio` to the logo's width-to-height ratio so it isn't distorted on resize (Excel's "keep aspect ratio" setting is unreliable, so this is set explicitly):

```vba
Public Const logoAspectRatio As Double = 1   ' logo width = aspectRatio × height
```

For a 300×100 logo use `3.0`; for a 200×200 square mark use `1.0`.

Also tune `logoHeightScale` — the logo height as a fraction of chart height:

```vba
Public Const logoHeightScale As Double = 0.1   ' 10% of chartHeight
```

Wide horizontal lockups often need `0.07`–`0.09`; square marks can stay at `0.1` or slightly larger.

### Step 4 — Test the logo

Rebuild the `.xlam`, apply any chart, and inspect the bottom-right corner. If the logo is positioned incorrectly, adjust `logoMarginRightProp` and `logoMarginBottomProp` in `modConfig.bas`.
