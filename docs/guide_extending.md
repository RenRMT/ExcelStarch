# Extending the Add-in with New Chart Types

The add-in is designed so that adding a chart type is localised to a few places: a new `.bas` module, a `ChartDefaults` factory in `modConfigCharts.bas`, one line in `modRibbonHandlers.bas`, one button in `CustomUI14.xml`, and an icon image. Nothing else needs to change.

---

## How chart modules work

Every chart type module follows the same two-tier pattern:

```vba
' === Private implementation ===
Private Sub BuildXxxChart()
    On Error GoTo CleanFail
    AppFast                          ' suppress screen updating (modAppState)

    Dim cht As Chart

    ' 1. Get a chart to style (re-style existing, or new from selection)
    Set cht = GetTargetChart(xlXxxChartType)
    If cht Is Nothing Then GoTo CleanExit

    ' 2. Run the shared formatting pipeline.
    '    Signature: ApplyChartPipeline cht, colorMode, defaults
    '    colorMode: "FILL" for bars/columns/area/pie, "LINE" for line/scatter
    '    defaults:  a ChartDefaults struct from modConfigCharts
    ApplyChartPipeline cht, "FILL", XxxChartDefaults()

    ' 3. Apply chart-type-specific property overrides
    cht.ChartGroups(1).GapWidth = seriesGapWidth
    ' ...
CleanExit:
    AppRestore                       ' always restore screen updating
    Exit Sub
CleanFail:
    AppRestore
    MsgError "BuildXxxChart"
End Sub

' === Public entry point (called by ribbon handler) ===
Sub XxxChart()
    BuildXxxChart
End Sub
```

The `Private/Public` split is intentional: `BuildXxxChart` holds all logic and is unreachable from outside the module. `XxxChart` is a one-line public stub that the ribbon handler and other modules (such as `modChartLollipop`, which calls `BarChart`) can call by name.

> **Application state:** wrap heavy builders with `AppFast` / `AppRestore` from `modAppState` and restore state on **both** the normal exit and the error handler, so screen updating is never left disabled. These calls are re-entrancy safe (depth-counted), so a builder that calls another builder works correctly.

---

## Step 1 — Create the module file

Create a new file named `modChartXxx.bas` (replace `Xxx` with the chart type name) using the following template:

```vba
Attribute VB_Name = "modChartXxx"
'==== Module: modChartXxx ====
' Brief description of this chart type and any notable behaviour.
'
' Variants (if applicable)
' --------
'   XxxChart        — xlXxxClustered: ...
'   StackedXxxChart — xlXxxStacked:   ...
Option Explicit


Private Sub BuildXxxChart()
    On Error GoTo CleanFail
    AppFast

    Dim cht As Chart

    Set cht = GetTargetChart(xlXxxClustered)   ' use the appropriate xlChartType constant
    If cht Is Nothing Then GoTo CleanExit

    ApplyChartPipeline cht, "FILL", XxxChartDefaults()   ' defaults from modConfigCharts
    Call RemoveShadow(cht)

    ' Chart-type-specific properties
    If cht.HasAxis(xlCategory) Then
        cht.Axes(xlCategory).MajorTickMark = xlTickMarkNone
        cht.Axes(xlCategory).MinorTickMark = xlTickMarkNone
    End If

    cht.ChartGroups(1).Overlap  = seriesOverlap
    cht.ChartGroups(1).GapWidth = seriesGapWidth
CleanExit:
    AppRestore
    Exit Sub
CleanFail:
    AppRestore
    MsgError "BuildXxxChart"
End Sub


Sub XxxChart()
    BuildXxxChart
End Sub
```

You will also need a `XxxChartDefaults()` factory in `modConfigCharts.bas` that returns a `ChartDefaults` struct describing which gridlines, axes, and legend this chart type should get. Copy an existing one (e.g. `ColumnChartDefaults`) as a starting point.

### Which `xlChartType` constant to use

Excel's chart type constants include:

| Constant | Chart shape |
|---|---|
| `xlBarClustered` | Horizontal bars, side by side |
| `xlBarStacked` | Horizontal bars, stacked |
| `xlBarStacked100` | Horizontal bars, 100% stacked |
| `xlColumnClustered` | Vertical bars, side by side |
| `xlColumnStacked` | Vertical bars, stacked |
| `xlColumnStacked100` | Vertical bars, 100% stacked |
| `xlLine` | Line chart |
| `xlLineMarkers` | Line chart with markers |
| `xlPie` | Pie chart |
| `xlDoughnut` | Donut chart |
| `xlXYScatter` | Scatter plot |
| `xlBubble` | Bubble plot |
| `xlArea` | Area chart |
| `xlAreaStacked` | Stacked area chart |
| `xlAreaStacked100` | 100% stacked area chart |

Use the constant that matches the chart *shape*, not the visual style. The pipeline handles the visual style.

### When to use "LINE" vs "FILL" pipeline mode

Pass `"FILL"` for charts where series are represented as filled areas (bars, columns, area, pie). Pass `"LINE"` for charts where series are represented as lines (line charts, scatter). This controls how `FormatSeriesColors` applies colour — fill colour vs line colour.

---

## Step 2 — Handle charts that cannot use the full pipeline

The 8-step pipeline in `ApplyChartPipeline` assumes the chart has a category axis, a value axis, and gridlines. Some chart types do not, and calling the pipeline on them raises errors. Three opt-out patterns exist, in increasing distance from the classic pipeline: a **partial pipeline** (pie/donut — same in-chart chrome, fewer steps), **composition** (lollipop — build an existing type then transform it), and a fully **separate pipeline** (chartex types — chrome cannot live in the chart at all).

### Pie/donut pattern: partial pipeline

`modChartPie.bas` demonstrates the opt-out pattern. Instead of calling `ApplyChartPipeline`, it calls only the pipeline steps that are safe for a pie chart:

```vba
Private Sub BuildPieChartWithDefaults(cht As Chart, ByRef defaults As ChartDefaults)
    InsertSource cht
    SetRoundChartSizeAndTitle cht, defaults     ' pie-specific sizing + title
    InsertLogo cht                              ' after sizing, so logo scales correctly
    ' Skips the axis/gridline pipeline steps entirely
    ApplySliceColors cht, cht.SeriesCollection(1).Points.Count
End Sub
```

Use this approach if your chart type lacks axes, has a non-standard layout, or needs a fundamentally different element arrangement. Call `InsertLogo` and `InsertSource` from `modChartBuilder` directly — they are both `Public` and safe to call in isolation.

### Composition pattern: call an existing chart, then post-process

`modChartLollipop.bas` demonstrates composing on top of an existing chart type:

```vba
Private Sub BuildLollipopChart()
    BarChart    ' run the full bar chart pipeline on a new chart
    Set cht = ResolveActiveChart()
    If cht Is Nothing Then Exit Sub

    ' Post-process: convert bars to lollipop sticks and dots
    For i = 1 To cht.SeriesCollection.Count
        ' hide bar fill, add error bars, apply arrowhead...
    Next i
End Sub
```

This pattern is appropriate for chart styles that are visual transforms of an existing type rather than entirely distinct chart types.

### ChartEx pattern: separate worksheet-chrome pipeline

The Excel 2016+ "chartex" chart types — **treemap, sunburst, waterfall, funnel, box & whisker, histogram** — are a different object family from the classic charts above. They are not just axis-less; they are stored in the file under a different schema (`cx:` chartex) that has **no `userShapes` slot**. The practical consequence for this add-in is decisive:

> `cht.Shapes.AddTextbox` and `cht.Shapes.AddPicture` raise **error 1004** on a chartex chart. A treemap (etc.) therefore **cannot own** the title/subtitle/figure/source/logo overlay boxes that every classic chart carries inside `cht.Shapes`.

Because the chrome cannot live inside the chart, it is built as **grouped worksheet shapes** instead. `modTreemapChrome.bas` is the reference implementation. This is a genuinely **separate pipeline** from the classic in-chart one in `modChartBuilder` — not a partial pipeline or a composition. The two do not share code paths (only the geometry constants in `modConfig` and the logo decode in `modEmbeddedImages` are reused).

```vba
' modChartPie.BuildTreemapChartWithDefaults (abridged)
If TypeName(cht.Parent) <> "ChartObject" Then MsgTreemapNeedsEmbedded: Exit Sub  ' embedded only
With cht.Parent: .Width = chartWidth: .Height = chartHeight: End With            ' fix the canvas
If pointscount > 0 Then ApplySliceColors cht, pointscount, silent:=True          ' tiles, not series
BuildTreemapChrome cht   ' worksheet shapes + group (modTreemapChrome) — NOT InsertSource/FormatTitle/InsertLogo
```

**Implications you must account for when adding another chartex type.** Model the new type on `modTreemapChrome` rather than on the classic builders, and understand that the separate pipeline reaches into several subsystems that assume in-chart chrome:

- **Coordinate model.** Classic chrome uses chart-relative coordinates (origin = chart top-left). Worksheet shapes use absolute sheet coordinates, so every position is offset by `cht.Parent.Left`/`.Top`. The chart is forced to `chartWidth × chartHeight` first so the same `modConfig` geometry constants apply.
- **Grouping is the contract.** The chrome shapes plus the `ChartObject` are grouped (`ws.Shapes.Range(names).Group`, named `ESTreemapGroup_<chartname>`) so they move and export as one unit. Pass the member-name list as a **`Variant` array**, not `String()` — `Shapes.Range` raises type-mismatch otherwise.
- **Re-run / re-style.** Re-running must ungroup, delete the prior chrome by name, and rebuild — otherwise chrome duplicates. The classic `SafeDeleteShape` only searches `cht.Shapes` and will **not** find worksheet chrome; use a worksheet-targeted delete (see `RemoveExistingTreemapChrome` / `SafeDeleteSheetShape`).
- **Export.** `Chart.Export` / `ExportAsFixedFormat` capture only the chart, so they **omit** worksheet chrome. `modExport` detects a chartex group selection and rasterises the whole group to PNG via a temporary chart (`CopyPicture` → temp `ChartObject` → `Chart.Export`). This is **screen-resolution PNG only** — no SVG/PDF, and softer than the classic export.
- **Toggles / Label Last Point.** These assume chrome is inside the chart (e.g. `MoveYAxisLabelBox` loops `cht.Shapes`; `LabelLastPoint` duplicates the `ChartObject` and relies on chrome riding along). They are **not wired** for chartex chrome. Treemaps sidestep this because they have no axis/legend; another chartex type with axes (waterfall, box & whisker) would need these tools taught about the worksheet group, or excluded from them.
- **Embedded charts only.** Chart sheets (`cht.Parent` is the `Workbook`) have no host worksheet for the shapes, and are rejected with `MsgTreemapNeedsEmbedded`.
- **Known prototype limitations.** Moving the chart after creation leaves chrome behind once the group is ungrouped; re-styling requires selecting the chart, not the group. Document these for any new chartex type that inherits the pipeline.

> **Roadmap note.** When a second chartex type is added, generalise this path rather than copying it: rename `modTreemapChrome` → `modChartExChrome`, gate on an `IsChartExType(chartType)` helper, and isolate per-type differences (has-axes? has-legend? tile vs. series colouring) behind small type-specific config — mirroring the `*ChartDefaults()` factory pattern. Keeping classic and chartex as two pipelines is deliberate: the worksheet+group approach is pure upside for chartex (which has no working in-chart path) but would be a regression for classic charts (which export crisply via `Chart.Export`).

---

## Step 3 — Register the ribbon handler

Open `modRibbonHandlers.bas` and add one line in the appropriate section:

```vba
'=== Chart creation ===
Public Sub Bar_onAction(control As IRibbonControl): BarChart: End Sub
Public Sub StackedBar_onAction(control As IRibbonControl): StackedBarChart: End Sub
Public Sub Lollipop_onAction(control As IRibbonControl): LollipopChart: End Sub
Public Sub Xxx_onAction(control As IRibbonControl): XxxChart: End Sub    ' ← add this
```

The naming convention is `<ButtonId>_onAction`. The `control As IRibbonControl` parameter is required by Excel's ribbon callback signature; you do not need to use it unless the handler needs to read `control.Tag`.

---

## Step 4 — Add the ribbon control in `CustomUI14.xml`

The chart groups are organised as **split buttons** (Column, Bar, Line & Area, Pie, Scatter, Complex). Each group is a `<splitButton size="large">` whose main `<button>` applies the group's default chart, and whose `<menu>` lists each variant. The **Complex** group holds the chartex types (treemap today); it is kept separate because those charts are not interchangeable with the classic types and use the separate pipeline described in Step 2.

**To add a variant to an existing group**, add a `<button>` inside that group's `<menu>`:

```xml
<menu id="ColumnMenu">
    <button id="ColumnMenuPlain"   image="i_chart_vbar"         label="Column Chart"             onAction="Column_onAction"/>
    <button id="ColumnMenuStacked" image="i_chart_stacked_vbar" label="Stacked Column Chart"     onAction="StackedColumn_onAction"/>
    <button id="ColumnMenu100"     image="i_chart_stacked_vbar" label="100% Stacked Column Chart" onAction="Stacked100Column_onAction"/>
    <button id="ColumnMenuXxx"     image="i_chart_xxx"          label="Xxx Column Chart"          onAction="Xxx_onAction"/>   <!-- ← add this -->
</menu>
```

Attributes:
- **`id`** — unique identifier for this control in the XML (every `id` must be unique across the whole file).
- **`image`** — the image ID registered in the Custom UI Editor. Must match the icon filename (without extension) you add in step 5.
- **`label`** — menu-item text.
- **`onAction`** — the VBA sub name in `modRibbonHandlers.bas`.

Menu items do **not** take a `size` attribute — only the parent `<splitButton size="large">` sets the size, which controls the large main button. Setting `size` on a `<splitButton>`'s child button or menu is not allowed.

### Adding a new chart group

If the new chart type deserves its own group, add a `<group>` with a large split button, modelled on the existing groups. The main button is the default; the menu holds the variants:

```xml
<group id="XxxGroup" label="Xxx">
    <splitButton id="XxxSplit" size="large">
        <button id="Xxx" image="i_chart_xxx" label="Xxx Chart"
                supertip="Style an Xxx chart following the COMPANY standards"
                onAction="Xxx_onAction"/>
        <menu id="XxxMenu">
            <button id="XxxMenuPlain"   image="i_chart_xxx"  label="Xxx Chart"     onAction="Xxx_onAction"/>
            <button id="XxxMenuVariant" image="i_chart_yyy"  label="Xxx Variant"   onAction="XxxVariant_onAction"/>
        </menu>
    </splitButton>
</group>
```

Groups appear left-to-right in the ribbon in the order they appear in the XML. The chart groups sit before the `FillColors` group.

---

## Step 5 — Add a button icon

Icon images must be square PNGs in the `icons/` folder. The ribbon uses 32×32 pixels for `size="large"` buttons and 16×16 for `size="normal"`. Provide a 32×32 PNG — Excel scales it if needed.

Name the file to match the `image` attribute: `i_chart_xxx.png`.

Add the image to the `.xlam` via the Office Custom UI Editor (File → Import → select the PNG, or right-click the image node in the tree).

---

## Step 6 — Rebuild and test

1. Import `modChartXxx.bas` into the VBE.
2. Add the handler line to `modRibbonHandlers.bas`.
3. Save the updated `.xlam`.
4. Embed the updated `CustomUI14.xml` and new icon via the Custom UI Editor.
5. Reload the add-in in Excel (via File → Options → Add-ins, untick and re-tick).
6. Test with:
   - A data range selected (Path 2: new chart from data)
   - An existing chart selected (Path 1: duplicate and restyle)
   - An empty chart (no series): should exit gracefully without an error

---

## Reference: modChartBar.bas as a template

The bar chart module is the canonical minimal example. It shows:
- The `Private/Public` two-tier pattern
- `GetTargetChart` call with the appropriate chart type constant
- `ApplyChartPipeline` with `"FILL"` mode
- `RemoveShadow` (used for clustered bar/column; not needed for stacked variants)
- Axis tick mark removal
- Gap width and overlap from `modConfig`

Use it as a copy-paste starting point and modify only the chart type constant and the axis/format properties that differ.

---

## Advanced: tag-based buttons (no new handler needed)

If the new chart type can be parameterised from the ribbon XML — as fill colours and ramps are — you can reuse an existing `onAction` handler by encoding the variant in the button's `tag` attribute. For example, a hypothetical "area with opacity" variant could reuse `Format_onAction` with a custom tag. Study `modFormatFill.bas` `ApplyFillFromTag` to understand the tag parsing pattern before using this approach.
