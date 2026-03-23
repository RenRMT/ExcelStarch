# Building the .xlam Add-in from Source Files

This guide walks through assembling the Excel add-in from the source files in this repository into a working `.xlam` file.

---

## What you need

| Tool | Purpose |
|---|---|
| Microsoft Excel (Windows) | The runtime and development environment |
| Visual Basic Editor (VBE) | Built into Excel — imports and compiles `.bas` modules |
| Office RibbonX Editor | Free tool to embed `CustomUI14.xml` into the `.xlam` binary |

The **Office RibbonX Editor** is a standalone and open-source tool available at [https://github.com/fernandreu/office-ribbonx-editor](https://github.com/fernandreu/office-ribbonx-editor). It reads and writes the Custom UI XML that defines the ribbon, which Excel does not expose through the VBE. Install it before starting.

---

## Step 1 — Create a new workbook and enable the Developer tab

1. Open Excel and create a new blank workbook.
2. If the **Developer** tab is not visible: File → Options → Customize Ribbon → tick **Developer** → OK.
3. Save the workbook as a macro-enabled file first:
   - File → Save As → choose **Excel Macro-Enabled Workbook (*.xlsm)** → name it `chart_styles` → Save.

   (You will convert to `.xlam` at the end. Starting as `.xlsm` lets you test before committing to the add-in format.)

---

## Step 2 — Import all VBA modules

1. Press **Alt + F11** to open the Visual Basic Editor.
2. In the Project Explorer (Ctrl+R to show it), you should see `VBAProject (chart_styles.xlsm)`.
3. Right-click the project name → **Import File...**
4. Navigate to the `modules/` folder of this repository.
5. Import the modules via File → Import File (or right-click → Import File in the Project Explorer).

6. After importing, verify the Project Explorer shows all 16 modules under **Modules**.

---

## Step 3 — Verify compilation

In the VBE menu: **Debug → Compile VBAProject**.

If this completes without an error dialog, all modules compiled successfully. If it raises an error, the first unresolved symbol will be highlighted — this usually means a module was imported out of order or a file is missing.

> **Common issue:** Excel's `IRibbonControl` type (used in `modRibbonHandlers.bas`) requires a reference to the Microsoft Office Object Library. Go to Tools → References and confirm **Microsoft Office 16.0 Object Library** (or the version installed on your machine) is ticked.

---

## Step 4 — Save as .xlam

1. In Excel (not the VBE), go to File → Save As.
2. Change the file type to **Excel Add-in (*.xlam)**.
3. Excel will suggest saving to `%AppData%\Microsoft\AddIns\`. You can save there, or save to the repository folder and load it manually. For development, keeping it in the repo folder is easier to track.
4. Name it `chart_styles.xlam` → Save.

The workbook is now an add-in. Excel will close the visible workbook — the add-in runs as a hidden workbook.

---

## Step 5 — Embed the ribbon XML with Office RibbonX Editor

The `.xlam` file is a ZIP archive internally. The Custom UI ribbon definition (`CustomUI14.xml`) must be embedded in this archive. Excel has no built-in interface for this — the Office RibbonX Editor does it.

1. Open **Office Ribbonx Editor**.
2. File → Open → navigate to your `.xlam` file.
3. In the left panel, right-click the file name → **Insert Office 2010+ Custom UI Part**.
   - This creates a `customUI14.xml` node (targeting the newer ribbon API used by Excel 2010 and later, which matches the `xmlns` declaration in this project's XML).
4. Select the `customUI14` node in the tree. Paste the entire contents of `CustomUI14.xml` from the repository into the editor panel on the right.
5. **Important:** Copy the icon image files to the RibbonX Editor:
   - File → **Import Custom UI** part images. Select all `.png` files from the `icons/` folder of the repository.
   - Alternatively, after saving, the editor provides a button to add image parts. Each image `id` in the XML (e.g., `i_chart_vbar`) must match an image file added here.
6. File → **Validate** — this checks the XML is well-formed and all image references resolve.
7. File → **Save** — this writes the XML into the `.xlam` archive.

---

## Step 6 — Load the add-in in Excel

1. In Excel: File → Options → Add-ins → at the bottom, **Manage: Excel Add-ins** → Go.
2. Click **Browse** → navigate to your `.xlam` → OK.
3. The add-in should now appear with a tick. Click OK.
4. Check the ribbon for the new tab (labeled whatever `orgName` is set to in `modConfig.bas`, e.g., "COMPANY Chart Styles").

If the tab does not appear, the ribbon XML was not embedded correctly. Re-open the `.xlam` in the RibbonX Editor and re-check step 5.

---

## Step 7 — Test a chart button

1. In a worksheet, enter a few rows of data.
2. Select the data range.
3. Click a chart button on the add-in tab (e.g., Bar Chart).
4. A styled chart should appear as a new chart object on the sheet.

If a VBA error appears, check the Immediate Window in the VBE (Ctrl+G) and the error message for the function name.

---

## Rebuilding after source changes

When you edit `.bas` files and want to rebuild:

1. Open the `.xlam` in the VBE (File → Open in the VBE, or load the `.xlam` via Add-ins and then open VBE).
2. In the Project Explorer, right-click the module to replace → **Remove** (choose No when asked to export).
3. Re-import the updated `.bas` file.
4. Debug → Compile to verify.
5. Save the `.xlam` (Ctrl+S in the VBE, or File → Save in Excel).

For the ribbon XML, repeat step 5 of the original build if `CustomUI14.xml` changed.

---

## Troubleshooting

| Symptom | Likely cause | Fix |
|---|---|---|
| Ribbon tab missing | CustomUI XML not embedded | Re-embed with RibbonX Editor |
| Compile error on `IRibbonControl` | Missing reference | Tools → References → tick Office library |
| Logo not appearing | Base64 string corrupted | Re-import `modEmbeddedImages.bas` |
| Icons not showing on ribbon buttons | Image files not imported | Re-add images via RibbonX Editor |
