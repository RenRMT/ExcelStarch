# Recommended Development Workflow

Working on an Excel add-in presents a specific challenge: the `.xlam` file is a binary ZIP archive that cannot be meaningfully diffed or merged in git. The VBA source, however, can be exported as plain-text `.bas` files that are fully trackable. This guide describes a workflow that keeps the source in git while managing the binary cleanly.

---

## The core problem

Excel stores VBA in a binary format inside the `.xlam` archive. You cannot edit `.bas` files on disk and have the change automatically reflected in the loaded add-in — you must import the file into the VBE, then save the `.xlam`. This means there are always two representations of the code:

1. **The `.bas` files** — the source of truth in git
2. **The `.xlam` binary** — the compiled, running artefact

The workflow below keeps these in sync without letting them drift.

---

## Tooling you need on every development machine

| Tool | Purpose |
|---|---|
| Excel (Windows) | Development and test environment |
| Visual Basic Editor (Alt+F11) | Edit and compile VBA modules |
| Office Custom UI Editor | Embed/update ribbon XML in the `.xlam` |
| Git | Version control for `.bas` and `.xml` source files |
| A text editor (VS Code, Notepad++) | Edit `.bas` and `.xml` files outside Excel |

---

## Project layout

```
chart_macro/
├── modules/                  ← All .bas source files (tracked in git)
│   ├── modConfig.bas
│   ├── modConfigColors.bas
│   ├── modMessages.bas
│   ├── modEmbeddedImages.bas
│   ├── modChartBuilder.bas
│   ├── modFormatSeries.bas
│   ├── modFormatFill.bas
│   ├── modRamp.bas
│   ├── modChartBar.bas
│   ├── modChartColumn.bas
│   ├── modChartLine.bas
│   ├── modChartPie.bas
│   ├── modChartLollipop.bas
│   ├── modChartTools.bas
│   ├── modExport.bas
│   └── modRibbonHandlers.bas
├── icons/                    ← Ribbon button icons (tracked in git)
├── CustomUI14.xml            ← Ribbon definition (tracked in git)
├── chart_styles.xlam         ← Binary artefact (NOT tracked in git)
└── .gitignore                ← Contains *.xlam
```

The `.xlam` should appear in `.gitignore`. Collaborators build their own local binary from the source.

---

## The edit → import → test → export loop

### Editing a module

1. **Edit the `.bas` file** in a text editor on disk (VS Code, Notepad++, etc.).
   - All the usual advantages apply: syntax highlighting, find-and-replace, side-by-side diffs.
   - Do not open and edit the VBE copy at the same time — one will overwrite the other.

2. **Import the file into the VBE:**
   - Open the `.xlam` in the VBE: File → Open → navigate to `chart_styles.xlam`.
   - In the Project Explorer, right-click the module you want to replace → **Remove Module** → choose **No** when asked to export (you already have the updated copy on disk).
   - File → **Import File** → select the updated `.bas`.

3. **Compile:** Debug → Compile VBAProject. Fix any errors before testing.

4. **Test** in Excel (see the Testing section below).

5. **Save the `.xlam`:** Ctrl+S in the VBE, or File → Save in Excel.

6. **Export the module back to disk** to capture any VBE-side edits (see Exporting below).

7. **Commit the `.bas` files** — never commit the `.xlam`.

### Exporting modules from the VBE

If you edited code directly in the VBE (which happens during debugging), export the module to keep the `.bas` file in sync before committing:

In the Project Explorer, right-click the module → **Export File** → navigate to `modules/` → overwrite the existing `.bas`.

> **Habit:** Always export before committing. A quick way to check: `git diff modules/` before staging. If you made a VBE change that isn't reflected in the diff, you forgot to export.

### Exporting all modules at once (VBE macro)

Add this utility sub to a scratch module in the VBE — do not commit it in the production modules. Run it from the Immediate Window to export all modules in one pass:

```vba
Sub ExportAllModules()
    Dim comp As VBComponent
    Dim path As String
    path = "C:\rprojects\chart_macro\modules\"   ' adjust to your repo path
    For Each comp In ThisWorkbook.VBProject.VBComponents
        If comp.Type = vbext_ct_StdModule Then
            comp.Export path & comp.Name & ".bas"
        End If
    Next comp
    Debug.Print "Export complete: " & path
End Sub
```

Run it via the Immediate Window: type `ExportAllModules` and press Enter.

---

## Working on the ribbon XML

`CustomUI14.xml` is a plain text file tracked in git. Edit it in a text editor. After editing:

1. Open the `.xlam` in **Office Custom UI Editor**.
2. Select the `customUI14` node.
3. Paste (or File → Import) the updated XML.
4. File → Validate — fix any XML errors.
5. File → Save.
6. Reload the add-in in Excel (File → Options → Add-ins → untick and re-tick).

### Why the ribbon requires a reload

Excel reads the ribbon XML only when the add-in loads. Saving the `.xlam` does not trigger a reload. You must deactivate and reactivate the add-in (or close and reopen Excel) to see ribbon changes.

---

## Testing

### Basic smoke tests after any change

| Test | What it checks |
|---|---|
| Select a data range → apply Bar Chart | Path 2 of `GetTargetChart` (new chart from range) |
| Double-click an existing chart → apply Column Chart | Path 1 of `GetTargetChart` (duplicate and restyle) |
| Apply a chart with no data | All guards against empty `SeriesCollection` |
| Apply a chart with 8+ series | `FormatSeriesColors` fallback to `colorSilver` |
| Apply a colour ramp with 8 series | `MsgRampTooManySeries` guard |
| Click a fill colour button with nothing selected | `MsgSelectTarget` guard |
| Export a chart as PNG | `RunChartExport` dialog and file write |

### Using the Immediate Window for rapid iteration

The Immediate Window (Ctrl+G in the VBE) lets you call functions directly on the active workbook state without going through the ribbon:

```vba
' Quickly re-run the pipeline on the active chart
ApplyChartPipeline ActiveChart, "FILL"

' Check a constant value
?chartWidth

' Test a colour conversion
?RGB(0, 119, 187)

' Force a specific error path
FormatSeriesColors ActiveChart, "INVALID"
```

This is faster than clicking ribbon buttons for repetitive tests, especially when iterating on layout constants.

### Inspecting chart element positions

The Immediate Window is also useful for reading back the positions Excel actually assigned, to cross-check against the constants in `modConfig.bas`:

```vba
?ActiveChart.PlotArea.Top
?ActiveChart.PlotArea.Height
?ActiveChart.Legend.Top
```

If an element is not where you expect, compare the printed value to the relevant constant and adjust.

---

## Version control discipline

### What to track in git

```
✓ modules/*.bas       — all VBA source
✓ CustomUI14.xml      — ribbon definition
✓ icons/*.png         — ribbon button icons
✗ *.xlam              — binary artefact, built locally
✗ *.xlsm              — any test workbooks
```

### Commit granularity

Commit at the module level: one commit per logical change, touching only the modules that changed. Avoid "WIP" commits that mix unrelated changes — the `.bas` diff is the only audit trail for what changed in the add-in.

Good commit message examples:
```
modChartBar: remove shadow on stacked variant
modMessages: centralise all inline MsgBox calls
modConfig: update canvas dimensions for A4 print output
CustomUI14: add Waterfall chart button to Other Graphs group
```

### Branching strategy

For a solo developer or small team:
- `main` — always contains a buildable, tested state
- Feature branches for any non-trivial change
- Merge only after rebuilding the `.xlam` and running smoke tests

### Handling the `.xlam` across machines

Each developer builds their own `.xlam` from the source. Document the build steps (or point to `guide_building_xlam.md`) in the `README`. Do not email or share `.xlam` files between developers — use git pull + local rebuild instead.

---

## Common mistakes and how to avoid them

| Mistake | Consequence | Prevention |
|---|---|---|
| Editing in VBE, forgetting to export | VBE changes overwritten on next import | Export all before committing; use `ExportAllModules` |
| Committing the `.xlam` | Binary blob in git history, no diff possible | `.gitignore` — `*.xlam` |
| Importing modules in the wrong order | Compile error on `Undefined Sub or Function` | Follow the import order in `guide_building_xlam.md` |
| Forgetting to reload after XML change | Ribbon shows old layout | Always deactivate/reactivate after XML edits |
| Testing on a chart that is currently active in edit mode | Can mask Path 2 bugs | Test both entry paths each time |
| Making a change in both the text editor and VBE simultaneously | The last save wins; one set of changes is lost | Edit in only one place at a time |

---

## Larger-scale changes: a worked example

**Scenario:** Add a new chart type (waterfall chart).

1. Create `modules/modChartWaterfall.bas` in a text editor using the template in `guide_extending.md`.
2. Add `Waterfall_onAction` to `modRibbonHandlers.bas`.
3. Add a `<button>` to `CustomUI14.xml`.
4. Add `i_chart_waterfall.png` to `icons/`.
5. Open the VBE, import `modChartWaterfall.bas` and the updated `modRibbonHandlers.bas`.
6. Debug → Compile. Fix any issues.
7. Save the `.xlam`.
8. Embed the updated `CustomUI14.xml` and new icon via Custom UI Editor.
9. Reload the add-in. Test both entry paths.
10. Export all modules back to disk.
11. `git add modules/ icons/ CustomUI14.xml && git commit -m "Add waterfall chart type"`

Total files changed in git: 3 (one new `.bas`, two updated files). The binary is not committed.
