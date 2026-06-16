# Manual Test Steps

This add-in is mostly chart-drawing code that manipulates the Excel object model, so most of it cannot be unit-tested — it has to be built and looked at. This is the repeatable checklist for that verification. (The pure logic — WCAG contrast maths, ramp ordering, tag parsing — is covered automatically by `modTestHarness`; run `RunAllTests` in the Immediate Window for that.)

## How to use

1. In the VBE: **Debug > Compile VBAProject** — must be clean before anything else.
2. In the Immediate Window: run `RunAllTests` — expect all `PASS:` lines, no halts.
3. Work through the sections below against a live workbook. Each row is *preconditions → action → expected*.
4. Suggested fixtures:
   - **Fixture A** — a small table, 1 category column + **3** numeric series columns, ~5 rows.
   - **Fixture B** — same shape but **8** series (exercises ramps near limits, legend wrapping).
   - **Fixture S** — two numeric columns (X, Y) for scatter; three (X, Y, size) for bubble.

---

## Chart creation

Each is a split button: the body applies the default; the dropdown lists variants. Select the fixture range first.

| Action | Fixture | Expected |
|---|---|---|
| Column / Stacked Column / 100% Stacked Column | A | Correct column type; house font/colours; logo (bottom-right), source (bottom-left), title/subtitle boxes present; category-axis tick marks hidden |
| Bar / Stacked Bar / 100% Stacked Bar | A | As above, horizontal |
| Lollipop (Bar menu) | A | Bars hidden; horizontal stick + round "candy" per series in palette colours |
| Line | A | Line per series; axis starts flush on first point; outside tick marks |
| Area / 100% Area (Line & Area menu) | A | Stacked filled areas flush to edges; axis lines white |
| Pie / Donut (Pie menu) | A (1 series) | Square canvas, centred plot; per-slice palette colours |
| Treemap (Pie menu, Excel 2016+) | A (1 series) | Hierarchical tiles; per-tile palette colours; no legend. If Excel rejects per-point colouring, tiles keep default colours and **no error dialog** appears |
| Scatter / Bubble (Scatter menu) | S | X/Y markers (bubble: third column sizes the bubbles) |

Spot-check that re-running a builder while a chart is **selected/active** re-styles it rather than creating a duplicate.

---

## Auto Style group

| Action | Preconditions | Expected |
|---|---|---|
| Apply Chart Style | A pie chart selected | Square sizing + per-slice colours applied in place |
| Apply Chart Style | A bar/column/area chart selected | Gridlines + axis labels + per-series FILL colours |
| Apply Chart Style | A line/scatter chart selected | Axis labels + per-series LINE colours |
| Apply Chart Style | A chart it can't fully style (e.g. radar) selected | Safe chrome only (font/border/logo/source/title); no error, no wrong geometry |

---

## Fill Colors group

| Action | Preconditions | Expected |
|---|---|---|
| Data Colors (split body) | Chart selected | Applies the last-used fill colour |
| A specific colour (Ocean…White) | A single series selected | Only that series recolours |
| A specific colour | Chart selected, no series picked | All series take that colour |
| Colour Ramps → Ocean (A) | Fixture A chart (3 series) | Dark→light left-to-right within the hue |
| Colour Ramps → any | Fixture B chart (8 series) | "Too many series" message (max 7 single-hue); no change |
| Diverging Ramps → Ocean—Coral (A\|B) | Fixture A chart, **odd** series count | Grey centre series; dark→light→…→light→dark across the two hues |
| Diverging Ramps → any | chart with **16** series | "Too many series" message (max 15) |

---

## Fill Actions group

| Action | Preconditions | Expected |
|---|---|---|
| Toggle Palette Order | Coloured chart | Series colour order switches between the two arrangements |
| Invert Ramp | Coloured chart | Fill colour order reverses across all series |
| Label Last Point | Line chart with a legend | Chart duplicated; series-name labels on the last point; plot narrowed; legend removed on the copy |

---

## Toggles group

Each cycles state in place on the selected chart. Click repeatedly and confirm the full cycle, then that it wraps back to the start.

| Action | Cycle |
|---|---|
| Toggle Data Labels | none → outside end → inside centre → none (stacked types skip "outside"; pie/donut inside labels use per-slice contrast colour) |
| Toggle Gridlines | none → horizontal → vertical → both → none |
| Toggle Axis Labels | none → x-axis → y-axis → both → none |
| Toggle Legend | on ↔ off, plot area resizes to match; single-series / treemap → informational message, no change |

---

## Export group

| Action | Expected |
|---|---|
| Chart Export → PNG / GIF / JPG / BMP / SVG / PDF | File written immediately in the chosen format; remembered format pre-selected next time |
| Chart Export | **No** overwrite warning if the file exists — known behaviour; confirm filename before OK |
| Chart Export (on Mac) | Informational "not supported" message; no crash |

---

## Edge cases worth a periodic pass

- No chart selected, then click any tool button → graceful "select a chart" message (not an error).
- A chart with a single series → legend toggle reports "not applicable".
- Ramps/diverging at the exact limits (7 single, 15 diverging) → applied; one above → capped with a message.
