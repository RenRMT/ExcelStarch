# Chart Styles Excel Add-in

A custom Excel add-in that applies organisational chart style standards from a dedicated ribbon tab. Select a data range, click a chart type, and the add-in creates a fully formatted chart — correct colours, fonts, sizing, layout, and branding — ready for publication.

This add-in was inspired by the [Urban Institute Data Visualisation Style Guide Excel Add-in](https://medium.com/urban-institute/introducing-the-urban-institute-data-visualization-style-guides-open-source-excel-add-in-14dfdfa50ebb), created by Jonathan Schwabish.

## Core Functionality

### Chart Creation

The ribbon provides buttons for the chart types in active use. Clicking a button creates a new chart from the current selection and applies the full formatting pipeline automatically.

| | Button | Chart Type |
|---|---|---|
| <img src="icons/i_chart_vbar.png" height="28"> | Column Chart | Clustered vertical bar |
| <img src="icons/i_chart_stacked_col.png" height="28"> | Stacked Column | 100% or absolute stacked vertical bar |
| <img src="icons/i_chart_hbar.png" height="28"> | Bar Chart | Clustered horizontal bar |
| <img src="icons/i_chart_stacked_bar.png" height="28"> | Stacked Bar | 100% or absolute stacked horizontal bar |
| <img src="icons/i_chart_lollipop.png" height="28"> | Lollipop Chart | Horizontal lollipop (bar chart with error-bar sticks and dot markers) |
| <img src="icons/i_chart_line.png" height="28"> | Line Chart | Standard line chart |
| <img src="icons/i_chart_pie.png" height="28"> | Pie Chart | Pie chart |
| <img src="icons/i_chart_donut.png" height="28"> | Donut Chart | Donut chart |

Each chart is created by duplicating a raw chart object, so the original data selection is preserved. The pipeline then applies: outer chart area formatting, plot area dimensions, axis styling, gridlines, series colours, title and subtitle text boxes, y-axis label box, logo, and source/notes placeholder.

The lollipop chart is built on top of the bar chart pipeline. After standard formatting is applied, each series bar is hidden and replaced with a horizontal error bar (extending from the value back to zero) formatted with an oval arrowhead at the value end — the stick and candy respectively. Stick and dot share the same brand colour per series.

### Brand Colour Palette

Seven data colours form the core palette:

| | Name | Description |
|---|---|---|
| <img src="icons/i_fill_ocean.png" height="28"> | Ocean | Primary blue |
| <img src="icons/i_fill_coral.png" height="28"> | Coral | Warm orange-red |
| <img src="icons/i_fill_sky.png" height="28"> | Sky | Light blue |
| <img src="icons/i_fill_pine.png" height="28"> | Pine | Teal-green |
| <img src="icons/i_fill_gold.png" height="28"> | Gold | Yellow |
| <img src="icons/i_fill_rust.png" height="28"> | Rust | Dark burnt orange |
| <img src="icons/i_fill_lavender.png" height="28"> | Lavender | Soft purple |

<img src="icons/i_fill_silver.png" height="28"> Silver and <img src="icons/i_fill_white.png" height="28"> White are available as neutral fills. Series beyond seven fall back to Silver.

**Palette ordering** can be toggled between two arrangements via the *Toggle Palette Order* button:

- **Contasting:** Ocean → Coral → Sky → Pine → Gold → Rust → Lavender
- **Complementary:** Ocean → Lavender → Sky → Pine → Gold → Coral → Rust

### Colour Ramps

The *Colour Ramps* group applies single-hue sequential palettes to the series of the active chart. Each brand colour has a seven-step ramp from light to dark. Steps are assigned in spread order (5, 1, 3, 6, 2, 4, 7) so that charts with fewer series achieve maximum contrast rather than a compressed range of similar tones.

Available ramps: Ocean, Coral, Sky, Pine, Gold, Rust, Lavender.

A diverging ramp mode is also available (invoked via VBA): two ramps are placed symmetrically — dark-to-light on the left side, light-to-dark on the right — with an optional neutral grey centre for odd series counts. Supports up to 15 series.

The **Invert** button reverses the current fill colour assignment across all series without re-applying a ramp, preserving any custom arrangement.

### Fill Colours

Per-colour buttons in the *Fill Colors* group apply a brand colour as a solid fill to any selected shape or chart element. Transparency can be specified via the button tag in the ribbon XML. A no-fill action removes the fill entirely.

| | | | | | | | | |
|---|---|---|---|---|---|---|---|---|
| <img src="icons/i_fill_ocean.png" height="28"> | <img src="icons/i_fill_coral.png" height="28"> | <img src="icons/i_fill_sky.png" height="28"> | <img src="icons/i_fill_pine.png" height="28"> | <img src="icons/i_fill_gold.png" height="28"> | <img src="icons/i_fill_rust.png" height="28"> | <img src="icons/i_fill_lavender.png" height="28"> | <img src="icons/i_fill_silver.png" height="28"> | <img src="icons/i_fill_white.png" height="28"> |
| Ocean | Coral | Sky | Pine | Gold | Rust | Lavender | Silver | White |

### Chart Tools

| | Button | Effect |
|---|---|---|
| <img src="icons/i_menu_grey.png" height="28"> | Reset to Grey | Resets all chart series fills to Silver |
| <img src="icons/i_menu_labels.png" height="28"> | Label Last Point | Adds series name labels to the final data point on line charts and narrows the plot area to make room |
| <img src="icons/i_menu_gridlines.png" height="28"> | Toggle Gridlines | Cycles the active chart through four gridline states: none → horizontal → vertical → both |
| <img src="icons/i_menu_none.png" height="28">  | Remove Legend | Deletes the chart legend and resizes the plot area to standard web dimensions |

### Export

The <img src="icons/i_menu_export.png" height="20"> *Chart Export* button exports the active chart object as an image or PDF. Supported formats: PNG, GIF, JPG, BMP, SVG, PDF. The chosen format is remembered between sessions. The export does not warn before overwriting an existing file.

---

## Requirements

- **Excel version:** 2013 or later. The ribbon XML uses the `customUI14` schema (Excel 2010+), and the `InsertChartField` API used for data labels requires Excel 2013+.
- **Operating system:** Windows only. The export function uses `GetSetting`/`SaveSetting` (Windows registry), and several chart formatting APIs behave differently or are unavailable on Mac Excel.

---

## Project Structure

| File | Role |
|---|---|
| `CustomUI14.xml` | Ribbon definition — tab layout, groups, buttons, image references, and `onAction` callback names |
| `modRibbonHandlers.bas` | Single entry-point layer for all ribbon callbacks. Thin wrappers only — one line per button, calling into the relevant module |
| `modChartBuilder.bas` | Shared formatting pipeline (`ApplyChartPipeline`) and all individual pipeline steps: `OuterFormat`, `FormatXAxisTitle`, `InsertLogo`, `InsertSource`, `FormatTitle`, `FormatGridlines`, `FormatXAxis` |
| `modChartColumn.bas` | Column and stacked column chart creation (`ColumnChart`, `StackedColumnChart`) |
| `modChartBar.bas` | Bar and stacked bar chart creation (`BarChart`, `StackedBarChart`) |
| `modChartLine.bas` | Line chart creation |
| `modChartPie.bas` | Pie and donut chart creation (`PieChart`, `DonutChart`) — shared layout helpers, variants differ only in chart type |
| `modChartLollipop.bas` | Lollipop chart creation — wraps the bar chart pipeline, then replaces bars with error-bar sticks and oval arrowhead dots |
| `modFormatSeries.bas` | Brand palette application to chart series (fill and line modes); palette order toggle |
| `modFormatFill.bas` | Solid fill application and removal for selected shapes |
| `modRamp.bas` | Single-hue sequential ramps, diverging ramps, and ramp inversion |
| `modConfigColors.bas` | Brand colour and ramp step constants |
| `modConfig.bas` | Layout, font, sizing, and export constants |
| `modEmbeddedImages.bas` | Organisation logo encoded as a Base64 string; decoded to a temp file at runtime for insertion into charts. Ribbon button icons are embedded separately into the `.xlam` via the Custom UI Editor |
| `modChartTools.bas` | Post-creation chart utilities: label last point, toggle gridlines, remove legend and resize, reset to grey |
| `modExport.bas` | Chart export dialog and file-write logic |
| `modMessages.bas` | Shared error and status message strings |

---

## Installation

The repository stores the add-in as individual `.bas` source files and an XML ribbon definition. To build a working `.xlam`:

**Tools required:**
- Microsoft Excel (2013+, Windows)
- The [Office Custom UI Editor](https://github.com/fernandreu/office-ribbonx-editor) (free, open source) for embedding the ribbon XML and button images

**Steps:**

1. **Create a blank add-in.** Open Excel, press `Alt+F11` to open the VBA editor, then save the file as an Excel Add-in (`.xlam`) via *File → Save As*.

2. **Import the modules.** In the VBA editor, right-click the project and choose *Import File*. Import every `.bas` file in the `modules/` folder.

3. **Embed the ribbon XML.** Close Excel. Open the `.xlam` in the Custom UI Editor. Create a new `customUI14` part and paste the contents of `CustomUI14.xml`. Import each ribbon button image (referenced by `image="..."` attributes in the XML) using the editor's image import function. Save and close.

4. **Replace the logo.** In `modEmbeddedImages.bas`, replace the Base64 string in `LogoPNG_Base64()` with a Base64-encoded version of your organisation's logo (SVG or PNG). Online Base64 encoders work for this.

5. **Enable the add-in.** Open Excel, go to *File → Options → Add-ins → Manage: Excel Add-ins → Go*, click *Browse*, and select the `.xlam` file. The *Chart Styles* tab will appear in the ribbon.

---

## Customisation

To adapt the add-in for a different organisation, two files cover most changes:

**`modConfig.bas`** — identity and layout:
```vb
Public Const orgName As String = "COMPANY"          ' used in ribbon labels and export settings
Public Const orgSupportContact As String = "COMPANY IT"
Public Const fontPrimary As String = "Calibri"      ' body and axis font
Public Const chartWidth As Double = 456.48           ' canvas size in points (6.34")
Public Const chartHeight As Double = 456.48
```

**`modConfigColors.bas`** — brand colours and ramp steps:
```vb
Public Const colorOcean As Long = 12285696    ' replace with your brand RGB values
Public Const colorCoral As Long = 6719743
' ... and the corresponding rampOcean1–7, rampCoral1–7 etc.
```

Colour values are stored as Excel `Long` RGB. Note that Excel uses BGR byte order internally, so the RGB values shown in the comments in `modConfigColors.bas` are in standard R,G,B notation for readability — use `RGB(R, G, B)` in VBA to convert.

The **logo** is set by replacing the Base64 string in `modEmbeddedImages.bas`. The ribbon **button icons** are embedded directly in the `.xlam` via the Custom UI Editor and are referenced by the `image="..."` attribute in `CustomUI14.xml`.

---

## Known Limitations

- **7-series cap on colour ramps.** Single-hue ramps support a maximum of 7 series; diverging ramps support up to 15 (7 + grey centre + 7). Charts exceeding the limit will not have a ramp applied and will show a warning.
- **No overwrite warning on export.** The Chart Export function will silently overwrite an existing file at the chosen path.
- **Windows only.** The add-in will not work on Mac Excel due to use of `GetSetting`/`SaveSetting` (Windows registry) and chart formatting APIs that behave differently on Mac.
- **Series limit on palette.** The brand palette has 7 colours. An 8th or higher series will be formatted in Silver (neutral grey) rather than a brand colour.
