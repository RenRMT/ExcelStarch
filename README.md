# ExcelStarch to fix your Excel charts

Ever encountered a chart or graph in the wild that just looked...Excelly? Charts made in Excel don't *have* to look bad, the default settings are just not very good. This configurable Excel add-in applies better default settings or specific organisational chart style standards from a dedicated ribbon tab. Select a data range, click a chart type, and the add-in creates a fully formatted chart.

This add-in was inspired by the [Urban Institute Data Visualisation Style Guide Excel Add-in](https://medium.com/urban-institute/introducing-the-urban-institute-data-visualization-style-guides-open-source-excel-add-in-14dfdfa50ebb), created by Jonathan Schwabish.

## Features

Select a data range and click a chart type button. The add-in creates a formatted chart with correct fonts, colours, sizing, and branding applied automatically. Eight chart types are supported: column, stacked column, bar, stacked bar, lollipop, line, pie, and donut.

Post-creation tools cover colour ramps (single-hue and diverging), per-element fill colours, palette order toggle, ramp inversion, gridline cycling, legend removal, last-point labelling, and PNG/PDF export.

See [docs/guide_users.md](docs/guide_users.md) for full button reference and usage instructions.

## Requirements

- **Excel version:** 2013 or later. The ribbon XML uses the `customUI14` schema (Excel 2010+), and the `InsertChartField` API used for data labels requires Excel 2013+.
- **Operating system:** Windows only. The export function uses `GetSetting`/`SaveSetting` (Windows registry), and several chart formatting APIs behave differently or are unavailable on Mac Excel.

## Installation

The repository stores source as `.bas` modules and `CustomUI14.xml`. The `.xlam` binary is not tracked. See [docs/guide_building_xlam.md](docs/guide_building_xlam.md) for the full build procedure.

## Customisation

All brand-specific settings are isolated in `modConfig.bas` and `modConfigColors.bas` with an optional embedded logo om `modEmbeddedImages.bas`. The ribbon tab label is hardcoded in `CustomUI14.xml` — find-and-replace `COMPANY` to update it. See [docs/guide_branding.md](docs/guide_branding.md) for the step-by-step procedure and full constant reference.
