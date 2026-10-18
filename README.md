# rethink

**rethink** is a clean-room PowerPoint charting add-in and presentation generator for Windows, written in C++20.
It follows the *behaviour* of think-cell (charts that are driven by a datasheet and recalculate automatically) without using any of its code, identifiers or artwork.

It consists of three parts that share one core:

| Part | What it does |
|---|---|
| **PowerPoint add-in** (`bin\rethink-addin.dll`) | Adds a **rethink** ribbon tab. Insert a waterfall chart, edit its data in a built-in datasheet, drag ▲▼ handles on the bars to change values, see a blue selection glow on rethink charts. |
| **Command line tool** (`bin\rethink.exe`) | Builds presentations from `.ppttc` JSON data (the documented think-cell automation format): fills text fields, native PowerPoint charts and rethink charts in a template. |
| **Layout engine** (`src/engine`) | Pure C++ (no COM, no XML). Computes chart geometry, labels, waterfall arithmetic (`e` = computed total) and the data table below a chart. Both the add-in and the CLI render its output. |

> Status: work in progress (Windows 11, PowerPoint / Microsoft 365). Implemented: **waterfall** (build-up/build-down, four orientations) and **Gantt / timeline** (insert via *rethink › Diagramm einfügen*, `"type": "gantt"` in `.ppttc`); see [docs/charting-ledger.md](docs/charting-ledger.md) for the full feature list and its status, and [HANDOFF.md](HANDOFF.md) for how to continue development.

---

## Quick start (try it in 5 minutes)

### 1. Requirements
- Windows 10/11 (x64) with **PowerPoint** (Microsoft 365 / Office 2016+, 64-bit)
- **Visual Studio 2022 Build Tools** (MSVC + CMake) and **Windows SDK ≥ 10.0.22621**
- No other libraries: rethink uses only Windows APIs (OPC packaging, MSXML 6, `Windows.Data.Json`, URLMon).

### 2. Build
Open a *Developer Command Prompt for VS 2022* (or run `vcvars64.bat`) in the repository folder:

```bat
cmake -S . -B %TEMP%\rtk -G "NMake Makefiles" -DCMAKE_BUILD_TYPE=Release
cmake --build %TEMP%\rtk
mkdir bin
copy %TEMP%\rtk\rethink.exe bin\
copy %TEMP%\rtk\rethink-addin.dll bin\
```

A short build folder (`%TEMP%\rtk`) avoids MSBuild's 260-character path limit when the repository lives in a deep path.

### 3. Register the add-in (current user only, no admin rights)
Close PowerPoint, then in PowerShell:

```powershell
.\scripts\register.ps1      # registers bin\rethink-addin.dll under HKCU
```

Start PowerPoint → tab **rethink** appears. Remove it again with `.\scripts\unregister.ps1`.

### 4. Try it
1. **rethink › Diagramme › Wasserfall einfügen** inserts a waterfall chart with sample data and a data table.
2. Click the chart: a thin **blue glow** with a tiny “rethink chart” label appears, plus **▲▼ handles** on every bar and a **table icon** at the bottom right.
3. Click ▲ / ▼ on a bar (Shift = 10× step) or **drag** a handle up/down (Esc cancels; one undo step per drag): the value changes in steps matching the precision you entered (12 → 1, 12,5 → 0,1), connectors and totals (`e`) recalculate live.
4. Click the **table icon** (or double-click the chart): the built-in **datasheet** opens. Type values (`12,5`, `48%`, `e`, dates), paste from Excel with Ctrl+V, right-click for rows/columns, percent and cell colour. Every entry updates the chart immediately; just close the window when done.

### 5. Try the command line tool
```powershell
.\ausprobieren.ps1            # builds samples\sample.ppttc and opens the result in PowerPoint
bin\rethink.exe build tests\waterfall.ppttc -o tests\out\waterfall.pptx
```
(`tests\waterfall.ppttc` needs the template from `tests\make-waterfall-template.ps1`.)

---

## The add-in in detail

| Ribbon command | Effect |
|---|---|
| **Diagramme › Wasserfall einfügen** | Inserts a rethink waterfall (sample data, data table). |
| **Diagramme › Datenblatt öffnen** | Opens the built-in datasheet of the selected rethink chart (same as the table icon or a double-click). |
| Erzeugen › Aus .ppttc erzeugen… | Builds a presentation from a `.ppttc` file (like the CLI). |
| Elemente › Diagramm einfügen / Elementname setzen | Inserts a native chart / names shapes so `.ppttc` data can fill them. |

- **Selection overlay:** a separate transparent window above PowerPoint (the document is never modified). Only the handles and the table icon accept clicks; everything else passes through to PowerPoint.
- **In-place updates:** each shape of a chart is named after its role (`segment r1 c3`, `connector c2`, …). Value changes move/resize the existing shapes instead of redrawing them, so your own formatting (e.g. a recoloured bar) survives. A ▲ click takes ~35 ms.
- **Resizing** the chart with the mouse re-lays it out after you release the button.
- **Data** lives inside the chart (shape tag `RETHINK_MODEL`, JSON) and travels with the presentation.
- **Excel as an alternative:** right-click in the datasheet → *In Excel bearbeiten* opens the data in a private Excel instance; edits apply live and closing never asks to save.

## `.ppttc` data and the `"chart"` extension

The CLI reads the JSON format documented for think-cell automation (`template`, `data[]`, `name`, `table`, cells with `string` / `number` / `percentage` / `date`, optional `fill`). Errors name the JSON path, e.g. `[0].data[3].table[1][2]`.

Shapes are found by their **name** (selection pane) or alt text `rethink:<name>`. Text fields of think-cell templates are recognised as well.

rethink adds an optional **`"chart"`** object to draw its own charts into a placeholder shape of the same name:

```json
{
  "name": "Bridge",
  "chart": { "type": "waterfall", "datatable": true, "tableRows": [2] },
  "table": [
    [null, {"string":"EBIT 2024"}, {"string":"Volume"}, {"string":"Costs"}, {"string":"EBIT 2025"}],
    [{"string":"EBIT"}, {"number":120}, {"number":35}, {"number":-28}, {"string":"e"}],
    [{"string":"Margin"}, {"percentage":8.1}, null, null, {"percentage":8.7}]
  ]
}
```

| `chart` key | Meaning (default) |
|---|---|
| `type` | `"waterfall"` |
| `datatable` | data table below the chart, columns aligned with the bars (`false`) |
| `tableRows` | row indices of `table` shown only in the data table (`[]`) |
| `labels` / `categories` / `connectors` | segment labels, category labels, waterfall connectors (`true`) |
| `decimals` | fixed number of decimals (automatic 0–2) |
| `fontSize`, `gap` | font size in pt (12), gap relative to bar width (0.5) |

**Waterfall semantics:** `{"string":"e"}` is a computed segment – a (sub)total down to the baseline; in the first column it is computed backwards (build-down bridge). Several data rows give several segments per column; at most one `e` per column.

## Tests

```powershell
%TEMP%\rtk\engine_test.exe                         # layout engine + datasheet + input parsing (unit tests)
tests\run-tests.ps1 -Exe bin\rethink.exe           # CLI validation and error messages
tests\waterfall-addin-test.ps1                     # add-in end to end via PowerPoint automation (needs registration)
tests\addin-test.ps1                               # older add-in commands
%TEMP%\rtk\addin_selftest.exe %TEMP%\rtk\rethink-addin.dll   # add-in COM self test, no PowerPoint needed
```

PowerPoint automation tests open PowerPoint (and briefly Excel) on screen. Real mouse interaction (clicking/dragging handles, double-click, the look of the glow) cannot be automated and must be checked by hand.

## Repository layout

| Path | Content |
|---|---|
| `src/engine/` | layout engine: waterfall arithmetic, geometry, labels, data table |
| `src/` | CLI and shared core: `.ppttc` parsing, OPC/XML packages, text/chart filling, OOXML rendering of rethink charts, XLSX read/write, datasheet conversion, cell input parsing |
| `addin/` | PowerPoint COM add-in: ribbon, element rendering via the PowerPoint object model, selection overlay, events, built-in datasheet window, Excel link |
| `scripts/` | register / unregister the add-in (HKCU) |
| `tests/`, `samples/` | automated tests, sample data and templates |
| `docs/charting-ledger.md` | feature list (spec from the documented behaviour) with status |

## Clean-room note

rethink was written from the publicly documented `.ppttc` format and the behaviour described in the user documentation. No code was decompiled or copied; no think-cell identifiers, names, icons or branding are used. Reverse-engineering notes about the original product are intentionally **not** part of this repository.

Note: UI strings and most code comments are currently German; this README and the project documents are English.
