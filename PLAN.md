# Implementation Plan: ExcelStarch Bug Fixes and Feature Enhancements

## Overview

This plan addresses 3 bug fixes, 3 new features, and 2 infrastructure changes to the ExcelStarch chart-styling VBA add-in. All changes must preserve the existing chart pipeline contract and ribbon integration.

---

## Recommended Implementation Order

```
Phase 1: Standalone Bug Fixes (no dependencies)
  Item 2 — FormatSeriesColors subscript out of range    [1 file]
  Item 3 — Pie chart legend left alignment              [1 file]

Phase 2: Error Handling Fix (no dependencies)
  Item 4 — Toggle data labels unsupported type error    [1 file]

Phase 3: Infrastructure Constants (prerequisite for Phase 5)
  Item 7 — Default chart formatting constants           [1 file]

Phase 4: Color Ramp Order Fix (no dependencies)
  Item 1 — Color ramp button order                      [1 file]

Phase 5: Pipeline Update (depends on Item 7)
  Item 8 — In-place chart modification + defaults       [9 files]

Phase 6: New Features (can use shared helpers from Item 8)
  Item 5 — Toggle legend                                [4 files]
  Item 6 — Toggle axis labels                           [3 files]
```

---

## Item 1: Color Ramp Buttons Apply Color in Wrong Order

### Current Behavior

`BuildColorRamp` in `modRamp.bas` (line 100-112) assigns ramp steps in "spread order" `[5,2,3,6,1,4,7]`. For a chart with 3 series, the assigned steps are `[5,2,3]` — step 5 is medium-dark, step 2 is light, step 3 is medium-light. The visual result is an inconsistent, non-sequential progression.

The diverging ramp (`BuildDivergingRamp`) uses the same priority order to **select** steps, but then **sorts them numerically** before assigning (line 148-155). This produces a smooth gradient. The single-hue ramp does NOT sort after selection.

### Root Cause

The spread order `[5,2,3,6,1,4,7]` is a **selection priority** (which steps to use when you have fewer than 7 series) but should not be the **assignment order**. After selecting the steps, they should be sorted numerically (like the diverging ramp does) so the visual result is a sequential dark-to-light gradient.

### Assumption

Dark-to-light assignment (series 1 = darkest selected step, series N = lightest selected step), matching the left side of a diverging ramp.

### Files Affected

- `modules/modRamp.bas` — `BuildColorRamp`

### Implementation Steps

1. After selecting the first `n` steps from the priority order, add the same bubble-sort that `BuildDivergingRamp` uses (ascending).
2. Then iterate in **descending** order (`For i = n To 1 Step -1`) to assign colors, producing dark-to-light.

### Edge Cases

- **1 series**: Only step 5 selected. No sorting needed. Produces mid-tone ramp step.
- **7 series**: All steps selected, sorted ascending, assigned descending = `[7,6,5,4,3,2,1]`. Series 1 = darkest, series 7 = lightest.
- **InvertColorRamp**: Still works — it snapshots current fills and reverses. No change needed.

### Dependencies

None.

---

## Item 2: FormatSeriesColors Subscript Out of Range for Series > 7

### Current Behavior

`GetPaletteColor` in `modFormatSeries.bas` (line 9-23) declares `palette(1 To 7)` and uses `IIf(i >= 1 And i <= 7, palette(i), colorNeutral2)`.

### Root Cause

**VBA's `IIf` evaluates BOTH branches regardless of the condition.** When `i = 8`, VBA evaluates `palette(8)` even though the condition is `False`, raising "Subscript out of range" before `IIf` can return `colorNeutral2`.

### Files Affected

- `modules/modFormatSeries.bas` — `GetPaletteColor` (line 22)

### Implementation Steps

1. Replace `IIf` with an `If...Then...Else` block:
   ```vb
   If i >= 1 And i <= 7 Then
       GetPaletteColor = palette(i)
   Else
       GetPaletteColor = colorNeutral2
   End If
   ```

### Testing

Create a chart with 8+ series. Apply any chart style. Verify series 1-7 get palette colors, series 8+ get `colorNeutral2`. No runtime error.

### Dependencies

None.

---

## Item 3: Pie Chart Legend Alignment

### Current Behavior

In `SetRoundChartSizeAndTitle` (modChartPie.bas, lines 126-131), the legend is positioned at top with a `.Top` value, but no `.Left` is set. Excel defaults to centering the legend horizontally within the chart area. The non-pie pipeline in `OuterFormat` does set `cht.Legend.Left = legendLeftPad`.

### Files Affected

- `modules/modChartPie.bas` — `SetRoundChartSizeAndTitle`

### Implementation Steps

1. Add `cht.Legend.Left = legendLeftPad` after `cht.Legend.Position = xlLegendPositionTop` and before `cht.Legend.Select`.
2. Add `cht.Legend.Font.Color = legendFontColor` for consistency with the non-pie pipeline.

### Edge Cases

- **Donut charts**: Use the same `SetRoundChartSizeAndTitle` via variant toggle. Fix applies to both.
- **Treemap**: Deletes the legend entirely. Not affected.

### Dependencies

None.

---

## Item 4: Toggle Data Labels Error on Unsupported Label Types

### Current Behavior

`ToggleDataLabels` in `modChartTools.bas` (lines 526-548) applies label positions unconditionally. Chart types like pie, line, scatter, and area do not support all positions (`xlLabelPositionOutsideEnd`, `xlLabelPositionCenter`), causing runtime error 1004.

### Files Affected

- `modules/modChartTools.bas` — `ToggleDataLabels`

### Implementation Steps

1. **Create helper**: `Private Function TryApplyLabelPosition(srs As Series, ByVal pos As Long) As Boolean` — attempts to set position, returns `False` on error.
2. **Update "OUTSIDE" case**: Try `xlLabelPositionOutsideEnd`. If fails, try `xlLabelPositionCenter`. If both fail, leave labels at default position.
3. **Update "INSIDE" case**: Try `xlLabelPositionCenter`. If fails, try `xlLabelPositionOutsideEnd`. If both fail, leave labels at default position.
4. **Guard state detection**: If reading `DataLabels.Position` fails (line 489-498), treat `currentState` as `"OTHER"` so the cycle still advances.

### Edge Cases

- **Treemap**: May not support `ApplyDataLabels` at all. Wrap in helper.
- **Scatter**: Supports `xlLabelPositionRight` etc., not `OutsideEnd`/`Center`. Falls back to default.
- **Mixed chart types (combo)**: Per-series approach handles this.

### Dependencies

None.

---

## Item 5: Remove Legend → Toggle Legend

### Current Behavior

`RemoveLegendAndResize` in `modChartTools.bas` (lines 328-348) only removes the legend and resizes. No way to restore it. No informative message for single-series charts.

### Files Affected

- `modules/modChartTools.bas` — Replace `RemoveLegendAndResize` with `ToggleLegend`
- `modules/modRibbonHandlers.bas` — Rename handler
- `CustomUI14.xml` — Update button id, label, supertip, onAction
- `modules/modMessages.bas` — Add `MsgLegendNotApplicable`
- `modules/modConfig.bas` — Reuse existing constants for both states

### Implementation Steps

1. **Add message**: `Public Sub MsgLegendNotApplicable()` — "This chart has only 1 data series or does not support a legend."
2. **Replace `RemoveLegendAndResize` with `ToggleLegend`**:
   - Guard: no chart → `MsgNoActiveChart`, exit.
   - Guard: 1 series or no legend support → `MsgLegendNotApplicable`, exit.
   - If `hasLegend = True`: remove legend, resize to no-legend dimensions (`removelegend*` constants).
   - If `hasLegend = False`: set `hasLegend = True`, format legend (position top, left pad, font color/size), resize to with-legend dimensions.
3. **Rename public entry point** to `ToggleLegendButton`.
4. **Update ribbon handler** to `ToggleLegendButton_onAction`.
5. **Update ribbon XML** — button id, label "Toggle Legend", supertip.
6. **Handle legend support detection**: Wrap `cht.hasLegend = True` in error handling for chart types that don't support re-adding legends.

### Edge Cases

- **Pie charts**: Use different pipeline and different plot area constants. Toggle must detect pie/donut and apply `piePlotAreaSize_legend` / `piePlotAreaSize_noLegend` and pie-specific legend positioning.
- **Charts from `ApplyChartStyle`**: May have non-standard dimensions. Toggle applies canonical dimensions regardless.

### Dependencies

- Weak dependency on Item 7 (legend default constant). Can be implemented independently.

---

## Item 6: Create a Toggle Axis Labels Function

### Current Behavior

No function exists to toggle axis tick labels. `ToggleAxes` toggles axis **visibility** (the entire axis), not just the labels.

### Files Affected

- `modules/modChartTools.bas` — Add `ToggleAxisLabels`
- `modules/modRibbonHandlers.bas` — Add handler
- `CustomUI14.xml` — Add button to Customisation group

### Implementation Steps

1. **Create `ToggleAxisLabels`** in `modChartTools.bas`:
   - Cycle: None → X only → Y only → Both → None
   - Read state via `TickLabelPosition` (`xlTickLabelPositionNone` vs `xlTickLabelPositionNextToAxis`)
   - To hide: `cht.Axes(axisType).TickLabelPosition = xlTickLabelPositionNone`
   - To show: `cht.Axes(axisType).TickLabelPosition = xlTickLabelPositionNextToAxis`
   - Guard: only operate on axes that exist (`cht.HasAxis`)
2. **Apply styling when showing**: `TickLabels.Font.Size = axisFontSize`, `TickLabels.Font.Color` matching pipeline conventions.
3. **Handle missing axes**: If axis removed via `ToggleAxes`, treat as "not visible" for cycle detection, skip assignment.
4. **Add ribbon handler**: `ToggleAxisLabelsButton_onAction`
5. **Add ribbon button**: id `ToggleAxisLabelsButton`, label "Toggle Axis Labels", in Customisation group after Toggle Axis Lines.

### Edge Cases

- **Scatter charts**: Both axes are value axes. `HasAxis(xlCategory)` still works.
- **Pie/donut**: No axes. Graceful no-op or informative message.
- **Bar charts**: X/Y swapped visually but `xlCategory`/`xlValue` are consistent in Excel's model.

### Dependencies

- Weak dependency on Item 7 (default axis label constant). Can share `axisNone`/`axisX`/`axisY`/`axisBoth` constants.

---

## Item 7: Default Chart Formatting Constants in modConfig.bas

### Current Behavior

Pipeline hardcodes formatting decisions inline. No centralized constants for gridline/axis/label/legend defaults.

### Files Affected

- `modules/modConfig.bas` — Add new constants section

### Implementation Steps

1. **Define axis selection constants**:
   ```vb
   Public Const axisNone As Long = 0
   Public Const axisX As Long = 1
   Public Const axisY As Long = 2
   Public Const axisBoth As Long = 3
   ```

2. **Add default formatting constants**:
   ```vb
   '=== Default chart formatting ===
   Public Const defaultGridlines As Long = 0      ' axisNone
   Public Const defaultAxisDisplay As Long = 0    ' axisNone
   Public Const defaultAxisLines As Long = 0      ' axisNone
   Public Const defaultAxisLabels As Long = 0     ' axisNone
   Public Const defaultLegend As Boolean = False  ' False = None
   ```
   All set to "none" per spec.

### Risks

- Setting all defaults to "none" means new charts will have no gridlines, no axes, no axis lines, no labels, no legend. This is the desired behavior but a significant visual change from current output.
- Existing charts not affected (only new charts through the pipeline).

### Dependencies

Item 8 is the consumer of these constants. This item must be implemented first.

---

## Item 8: Update Chart Formatting Pipeline (In-Place Modification)

### Current Behavior

`GetTargetChart` (modChartBuilder.bas, lines 517-539) always duplicates the chart. Every formatting operation creates a duplicate. The request: modify existing charts in-place, only create new charts from range selection.

### Files Affected

- `modules/modChartBuilder.bas` — Modify `GetTargetChart`, update `ApplyChartPipeline`
- `modules/modChartBar.bas` — Add `HasAxis` guards
- `modules/modChartColumn.bas` — Add `HasAxis` guards
- `modules/modChartLine.bas` — Add `HasAxis` guards
- `modules/modChartArea.bas` — Add `HasAxis` guards
- `modules/modChartScatter.bas` — Add `HasAxis` guards
- `modules/modChartLollipop.bas` — Verify chain behavior
- `modules/modChartPie.bas` — Apply `defaultLegend` in pie pipeline

### Implementation Steps

1. **Modify `GetTargetChart` to not duplicate**:
   - If active chart exists: return it directly, set `chartType` to requested type.
   - If range selected: create new chart via `AddChart2`, return directly (no duplication).

2. **Handle chart type conversion**: Set `cht.chartType = chartType` on existing charts. Wrap in error handling for incompatible conversions.

3. **Add default formatting application to `ApplyChartPipeline`** (runs after existing 8 steps):
   - Apply `defaultGridlines` — add/remove gridlines per constant.
   - Apply `defaultAxisDisplay` — set `HasAxis` per constant.
   - Apply `defaultAxisLines` — set axis line visibility per constant.
   - Apply `defaultAxisLabels` — set `TickLabelPosition` per constant.
   - Apply `defaultLegend` — add/remove legend and resize.

4. **Update pie pipeline**: Add legend default handling to `SetRoundChartSizeAndTitle`.

5. **Verify lollipop chain**: `BuildLollipopChart` calls `BarChart` which calls `GetTargetChart`. After in-place change, `ActiveChart` after `BarChart` is the same chart. Verify this works.

6. **Add `HasAxis` guards** in all builder modules' post-pipeline code that touches axes.

### Critical Risk

**With `defaultAxisDisplay = axisNone`, the pipeline removes all axes. But chart-type-specific code that runs AFTER the pipeline (e.g., `modChartBar.bas` setting tick marks) will fail because the axis no longer exists.** The default-application step must run LAST, and all post-pipeline axis code must guard with `If cht.HasAxis(xlCategory) Then`.

### Other Risks

- Users lose the implicit "undo by deleting the duplicate" safety net.
- `AddChart2` without duplication may expose Excel default formatting bugs. Test carefully.
- `GrayOutChart` and `LabelLastPoint` manage their own duplication — not affected.

### Dependencies

- **Item 7** (REQUIRED): Default constants must exist.
- Items 5, 6 can share helpers for legend/axis label manipulation.

---

## Testing Strategy

### After Every Phase

- `Debug > Compile VBAProject` must pass.

### Phase 1 (Items 2, 3)

- Chart with 10 series → column builder → no error, series 8-10 are Steel.
- Pie chart with 5+ categories → legend is left-aligned.

### Phase 2 (Item 4)

- Toggle data labels on: bar, column, line, pie, donut, area, scatter, treemap.
- No runtime errors on any chart type.

### Phase 3 (Item 7)

- Compile only. Constants not consumed until Phase 5.

### Phase 4 (Item 1)

- Charts with 1, 3, 5, 7 series → apply each single-hue ramp.
- Series 1 = darkest, series N = lightest.
- InvertColorRamp reverses order. Diverging left side matches single-hue order.

### Phase 5 (Item 8)

- Select existing bar chart → click Column → chart converts in-place (no duplicate).
- Select range → click Line → single new chart created.
- All defaults "none" → no gridlines, axes, lines, labels, legend on new charts.
- Lollipop chain works in-place.
- No errors from axis-touching code in builders.
- GrayOutChart / LabelLastPoint still duplicate.

### Phase 6 (Items 5, 6)

- Toggle legend: multi-series removes/restores. Single-series shows message. Pie uses pie dimensions.
- Toggle axis labels: cycle None → X → Y → Both → None on bar chart. Removed axis gracefully skipped.

---

## Files Summary

| File | Items |
|---|---|
| `modules/modConfig.bas` | 7, 8 |
| `modules/modConfigColors.bas` | Reference only |
| `modules/modFormatSeries.bas` | 2 |
| `modules/modRamp.bas` | 1 |
| `modules/modChartPie.bas` | 3, 8 |
| `modules/modChartTools.bas` | 4, 5, 6 |
| `modules/modChartBuilder.bas` | 8 |
| `modules/modChartBar.bas` | 8 (HasAxis guards) |
| `modules/modChartColumn.bas` | 8 (HasAxis guards) |
| `modules/modChartLine.bas` | 8 (HasAxis guards) |
| `modules/modChartArea.bas` | 8 (HasAxis guards) |
| `modules/modChartScatter.bas` | 8 (HasAxis guards) |
| `modules/modChartLollipop.bas` | 8 (verify chain) |
| `modules/modRibbonHandlers.bas` | 5, 6 |
| `modules/modMessages.bas` | 5, 6 |
| `CustomUI14.xml` | 5, 6 |
