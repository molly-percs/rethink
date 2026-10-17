# Handoff – how to continue rethink

This file is for anyone picking up development. It describes the current state, how the pieces fit together and what to do next. Keep it updated with every push.

## Current state (2026-10-17)

- **Working and tested:** waterfall chart (build-up) with `e` totals and data table; CLI `.ppttc` filling; add-in with ribbon tab, insert waterfall, selection overlay (glow, ▲▼ step handles, table icon), built-in datasheet with live updates, optional Excel datasheet, in-place updates, relayout on resize.
- **Implemented, not yet compiled with MSVC or run in PowerPoint** (written in a Linux container: engine unit tests pass with g++, the add-in files pass a mingw-w64 syntax check with stubbed WRL/WinRT/MSXML headers; nothing else). First thing on Windows: build, run `engine_test`, `addin_selftest`, `tests\waterfall-addin-test.ps1` (new section 3b), then check by hand:
  - **Dragging the ▲▼ handles.** `addin/overlay.cpp` captures the mouse on a ▲▼ press; release without moving (< `SM_CYDRAG`) is a click, otherwise `Drag` gestures are posted, coalesced to the newest mouse position (this is the throttling: one render per message-loop turn), then exactly one `DragEnd` (button up) or `DragCancel` (Esc, capture lost). `addin/events.cpp` (`dragGesture`) freezes value, step (`stepDecimals`) and scale at the start: pixels per unit = `emuPerUnit(drawing)` (engine) × PowerPoint's points-to-pixels factor; all-zero charts fall back to 4 px per step. `addin/elements.cpp`: `beginSegmentDrag` freezes model, options and layout box and calls `Application.StartNewUndoEntry`; `dragSegmentTo` re-runs only the pure layout (no worker thread) and updates only prims that differ from the last drawing (`samePrim`), returning the new segment rectangles so the overlay is rebuilt without querying shapes; `endSegmentDrag` writes `RETHINK_MODEL`/`RETHINK_LAYOUT` once and syncs an open built-in datasheet and Excel cell. Errors during the drag are reported after the mouse is released. Automation: `DragSegment(row, column, "6;15;30")` runs the same `dragGesture` path and returns the frozen scale.
  - **▲▼ click step** now uses `stepDecimals` of the clicked value (45 → 46 even if another value has decimals).
  - **Glow follows the chart while PowerPoint moves it:** thread-local `WH_MOUSE` hook installed in `startEvents`. Press inside the element → after the drag threshold the overlay window is offset by the cursor delta (`offsetOverlay`); press on one of PowerPoint's 8 size grips or the rotation grip → overlay hidden until release. While the button is down the 150 ms tracking timer leaves the overlay alone; after release `placeOverlay` re-shows it because `overlayDisplaced()` is true.
  - Deviation from the earlier plan: the segment value is **not** stored in the `RETHINK_SEG` tag; the drag scale comes from the engine drawing at drag start, so the tag stays `pos`/`neg`/`e`.
- **Requested next:** insert menu like the original (popup gallery of all chart types; arrows around each icon pick the orientation; unimplemented types greyed out), build-down waterfall, chart orientation (up/down/left/right as an engine transform), then a Gantt chart slice.

See [docs/charting-ledger.md](docs/charting-ledger.md) for the full feature status.

## Build & run

```bat
:: Developer Command Prompt for VS 2022, in the repository folder
cmake -S . -B %TEMP%\rtk -G "NMake Makefiles" -DCMAKE_BUILD_TYPE=Release
cmake --build %TEMP%\rtk
```
- Copy `rethink.exe` and `rethink-addin.dll` into `bin\` (create it; it is git-ignored). **Close PowerPoint first** – it locks the registered DLL.
- `scripts\register.ps1` registers the add-in for the current user (HKCU only); `scripts\unregister.ps1` removes it.
- Use a short build directory: MSBuild fails on paths > 260 characters.

## Tests

| Command | Covers |
|---|---|
| `%TEMP%\rtk\engine_test.exe` | engine (waterfall arithmetic, layout bounds), datasheet XLSX round trip, cell input parsing, drag/step maths and scale |
| `%TEMP%\rtk\addin_selftest.exe %TEMP%\rtk\rethink-addin.dll` | COM class factory, ribbon XML, callbacks, no PowerPoint needed |
| `tests\run-tests.ps1 -Exe bin\rethink.exe` | CLI validation and error messages |
| `tests\waterfall-addin-test.ps1` | add-in end to end through PowerPoint automation (incl. drag via `DragSegment`, undo) |
| `tests\addin-test.ps1` | older add-in commands (24 checks) |

Automation cannot click or drag; every interactive feature has an automation entry point on `COMAddIns("rethink.addin").Object` (`InsertWaterfall`, `AdjustSegment`, `DragSegment`, `OverlayInfo`, `OpenDatasheet`, `DatasheetEnter`, `DatasheetPaste`, `DatasheetCell`, `CloseDatasheet`, `OpenExcelDatasheet`, `DatasheetWorkbook`, `ApplyDatasheets`, `EventStatus`, …) that runs the same code path. Add one for each new interaction.

## Architecture in one page

```
.ppttc JSON ──► src/ppttc.*  ──► DataItem {name, table, chartJson}
                                     │
                       src/engine/*  ▼  layoutChart(item, options, box) → Drawing {prims[]}
                                     │                     (pure C++, unit-tested)
              ┌──────────────────────┴───────────────────────┐
   src/render_element.* (OOXML group + RETHINK_MODEL tag)   addin/elements.* (PowerPoint object model,
   used by the CLI (src/build.cpp)                           role-named shapes, in-place update, tags)
                                                                   │
                       addin/events.* (selection/double-click events, overlay tracking, handle actions)
                       addin/overlay.* (layered window: glow, label, handles, table icon)
                       addin/datasheet_window.* (built-in spreadsheet-like datasheet)
                       addin/addin.cpp (COM add-in, ribbon callbacks, automation interface)
```

Key rules:
- The engine never touches COM or XML; renderers never compute layout.
- Shapes inside a rethink chart are named by role (`segment r1 c3`, `label r1 c3`, `connector c1`, `table cell r1 c4`, …). Same roles ⇒ update in place; different roles ⇒ full redraw (keeping name, z-order).
- Group tags: `RETHINK_MODEL` (JSON data item), `RETHINK_LAYOUT` (layout box + group rect when last drawn). Child tags: `RETHINK_STYLE` (last colour set by rethink, to detect user formatting), `RETHINK_SEG` (segment kind).
- During a handle drag the overlay holds the mouse capture: never show a message box mid-drag (collect the error, report after `DragEnd`), and the prim cache (`samePrim` against the last drawing) is only valid for that reason.
- WinRT JSON, OPC and MSXML run on a worker thread (`onWorker`) – never initialise WinRT on PowerPoint's UI thread.
- Never redraw or select inside a mouse handler of the overlay; post a message and act afterwards. Never block on Excel; treat `RPC_E_CALL_REJECTED` / `RPC_E_SERVERCALL_RETRYLATER` as "retry next tick".

## Conventions

- C++20, MSVC `/W4 /permissive- /utf-8`, the add-in builds with `/WX`.
- UI strings are German (the current user's language); repository documents and commit messages are English.
- Clean-room: use the documented behaviour only; no decompiling, no think-cell names, identifiers, icons or artwork; reverse-engineering notes stay out of the repository.
