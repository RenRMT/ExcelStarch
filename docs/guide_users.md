# Using the Chart Styles Add-in

Once the add-in is installed, a **COMPANY Chart Styles** tab appears in the Excel ribbon. All buttons live in this tab, organised into groups.

---

## Creating a chart

1. Select a data range in any worksheet.
2. Click a chart type button. The add-in creates a formatted chart as a new object on the sheet.

If a chart is already active (double-clicked into edit mode) or selected (single-clicked), the button duplicates and reformats that chart instead of creating one from the selection.

| | Button | Chart Type |
|---|---|---|
| <img src="../icons/i_chart_vbar.png" height="28"> | Column Chart | Clustered vertical bar |
| <img src="../icons/i_chart_stacked_col.png" height="28"> | Stacked Column | 100% or absolute stacked vertical bar |
| <img src="../icons/i_chart_hbar.png" height="28"> | Bar Chart | Clustered horizontal bar |
| <img src="../icons/i_chart_stacked_bar.png" height="28"> | Stacked Bar | 100% or absolute stacked horizontal bar |
| <img src="../icons/i_chart_lollipop.png" height="28"> | Lollipop Chart | Horizontal lollipop (bar with error-bar sticks and dot markers) |
| <img src="../icons/i_chart_line.png" height="28"> | Line Chart | Standard line chart |
| <img src="../icons/i_chart_pie.png" height="28"> | Pie Chart | Pie chart (up to 5 slices) |
| <img src="../icons/i_chart_donut.png" height="28"> | Donut Chart | Donut chart (up to 5 slices) |

The pipeline applies automatically: chart size, font, axis styling, gridlines, series colours, title and subtitle text boxes, y-axis label, logo, and a source/notes placeholder.

### Editing placeholder text

After creating a chart, click into the text boxes to replace the placeholder text:

- **TitleBox** — "Title in 20pt sentence case"
- **SubTitleBox** — "Subtitle in 16pt sentence case"
- **YAxisLabelBox** — "Y axis title (unit)"
- **XAxisBox** — "X axis title (unit)"
- **SourceBox** — "Source: …" / "Notes: …"

---

## Colour palette

Seven data colours are used for multi-series charts, applied in palette order:

| | Name | Description |
|---|---|---|
| <img src="../icons/i_fill_ocean.png" height="28"> | Ocean | Primary blue |
| <img src="../icons/i_fill_coral.png" height="28"> | Coral | Warm orange-red |
| <img src="../icons/i_fill_sky.png" height="28"> | Sky | Light blue |
| <img src="../icons/i_fill_pine.png" height="28"> | Pine | Teal-green |
| <img src="../icons/i_fill_gold.png" height="28"> | Gold | Yellow |
| <img src="../icons/i_fill_rust.png" height="28"> | Rust | Dark burnt orange |
| <img src="../icons/i_fill_lavender.png" height="28"> | Lavender | Soft purple |

Silver and White are available as neutral fills. Any series beyond seven falls back to Silver.

### Palette order

The *Toggle Palette Order* button switches between two series colour arrangements:

- **Contrasting** (default): Ocean → Coral → Sky → Pine → Gold → Rust → Lavender
- **Complementary:** Ocean → Lavender → Sky → Pine → Gold → Coral → Rust

Toggle before applying a chart type, or re-apply the chart type after toggling.

---

## Fill colours

The *Fill Colors* group applies a solid colour fill to the selected chart element or shape. Select a series bar, a plot area, a text box, or any shape, then click a colour.

| | | | | | | | | |
|---|---|---|---|---|---|---|---|---|
| <img src="../icons/i_fill_ocean.png" height="28"> | <img src="../icons/i_fill_coral.png" height="28"> | <img src="../icons/i_fill_sky.png" height="28"> | <img src="../icons/i_fill_pine.png" height="28"> | <img src="../icons/i_fill_gold.png" height="28"> | <img src="../icons/i_fill_rust.png" height="28"> | <img src="../icons/i_fill_lavender.png" height="28"> | <img src="../icons/i_fill_silver.png" height="28"> | <img src="../icons/i_fill_white.png" height="28"> |
| Ocean | Coral | Sky | Pine | Gold | Rust | Lavender | Silver | White |

---

## Colour ramps

A colour ramp applies a single-hue sequential palette to all series of the active chart, ranging from light to dark. Select or activate a chart, then click a ramp button.

Steps are assigned in spread order (5, 1, 3, 6, 2, 4, 7) so that charts with fewer series still achieve maximum contrast.

| | Ramp | | Ramp |
|---|---|---|---|
| <img src="../icons/i_ramp_ocean.png" height="28"> | Ocean | <img src="../icons/i_ramp_gold.png" height="28"> | Gold |
| <img src="../icons/i_ramp_coral.png" height="28"> | Coral | <img src="../icons/i_ramp_rust.png" height="28"> | Rust |
| <img src="../icons/i_ramp_sky.png" height="28"> | Sky | <img src="../icons/i_ramp_lavender.png" height="28"> | Lavender |
| <img src="../icons/i_ramp_pine.png" height="28"> | Pine | | |

Maximum 7 series for single-hue ramps.

### Diverging ramps

A diverging ramp uses two hues: dark-to-light on the left side of the chart, light-to-dark on the right. For an odd number of series, the centre series is assigned a neutral grey.

| | Diverging ramp | Series limit |
|---|---|---|
| <img src="../icons/i_ramp_oceancoral.png" height="28"> | Ocean — Coral | 15 |
| <img src="../icons/i_ramp_oceangold.png" height="28"> | Ocean — Gold | 15 |
| <img src="../icons/i_ramp_oceanrust.png" height="28"> | Ocean — Rust | 15 |
| <img src="../icons/i_ramp_pinegold.png" height="28"> | Pine — Gold | 15 |
| <img src="../icons/i_ramp_pinelavender.png" height="28"> | Pine — Lavender | 15 |
| <img src="../icons/i_ramp_pinerust.png" height="28"> | Pine — Rust | 15 |

### Invert ramp

The <img src="../icons/i_menu_invert.png" height="20"> *Invert Ramp* button reverses the current fill colour order across all series without re-applying a ramp. Useful for flipping a ramp direction or reversing a custom colour arrangement.

---

## Chart tools

| | Button | Effect |
|---|---|---|
| <img src="../icons/i_menu_order.png" height="28"> | Toggle Palette Order | Switch series colour order between Contrasting and Complementary |
| <img src="../icons/i_menu_invert.png" height="28"> | Invert Ramp | Reverse fill colour order across all series |
| <img src="../icons/i_menu_grey.png" height="28"> | Reset to Grey | Reset all series fills to Silver |
| <img src="../icons/i_menu_labels.png" height="28"> | Label Last Point | Add series name labels to the last data point (line charts); narrows the plot area to make room |
| <img src="../icons/i_menu_gridlines.png" height="28"> | Toggle Gridlines | Cycle the active chart through four states: none → horizontal → vertical → both |
| <img src="../icons/i_menu_none.png" height="28"> | Remove Legend | Delete the chart legend and resize the plot area |

---

## Exporting a chart

1. Select or activate a chart.
2. Click <img src="../icons/i_menu_export.png" height="20"> *Chart Export*.
3. Choose a folder, file name, and format. Supported formats: **PNG, GIF, JPG, BMP, SVG, PDF**.
4. Click OK. The file is written immediately.

The chosen format is remembered between sessions. The export **does not warn before overwriting** an existing file — check the filename before confirming.

For higher-resolution images (e.g. for print), right-click the chart and select *Save as Picture* instead.
