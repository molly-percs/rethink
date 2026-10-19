# Handoff – how to continue rethink

This file is for anyone picking up development. It describes the current state, how the pieces fit together and what to do next. Keep it updated with every push.

## Current state (2026-10-19)

Built and tested on Windows 11 ARM64 (Surface, Snapdragon) with x64 PowerPoint (Microsoft 365, English UI) and Visual Studio Build Tools 2026 (MSVC 14.50), x64 target.

- **Waterfall – done:** build-up and build-down (the flipped icon of the original: `e` in the first column, computed backwards), four orientations (`"orientation": "up" | "down" | "right" | "left"`, from the insert gallery arrows), data table (below the chart, or as columns beside horizontal bars), `e` totals, connectors. Add-in: ▲▼ (◀▶ when horizontal) click and drag with frozen scale, one undo entry per drag, built-in and Excel datasheet, in-place updates, relayout on resize. Verified end to end in PowerPoint (`tests\waterfall-addin-test.ps1`, 44 checks).
- **Gantt / timeline – done (slice 4):** see "Gantt in one page" below. Verified with 235 engine checks, the CLI (`tests\gantt.ppttc`), 98 PowerPoint automation checks (`tests\gantt-addin-test.ps1`) and real mouse input (`tests\gantt-mouse-test.ps1`: grip drag, body drag, click selects the bar).
- **Insert gallery:** ribbon *Diagramm einfügen* opens a popup of all chart types (own GDI icons); unimplemented types are greyed out; arrows around orientable icons pick the orientation. Quick buttons *Wasserfall* and *Gantt / Zeitplan*.
- **Glow follows the chart while PowerPoint moves it** (mouse hook): implemented in the previous commit, still only checked by reading the code – check by hand.
- **Ribbon layout:** groups *Struktur*, *Elemente*, *Daten*, *Text* and five large menus (*Layout*, *Farbe*, *Qualität*, *Zusammenarbeit*, *Mehr*), arranged with explicit `<box>` columns. Own controls use `getImage="RethinkImage"`; the `tag` carries `icon:<name>` plus optionally `mso:<idMso>` (`RethinkMso` → `ExecuteMso`) or `cmd:<Gantt command>` (`RethinkCommand`); the eight icons listed in `isLargeIcon` (addin.cpp) are requested at 32 px, all others at 16 px. Not yet implemented items are static `enabled="false"` with a „geplant“ supertip. Every `idMso` was checked with `CommandBars.GetLabelMso`.
- **Ribbon:** PowerPoint accepts the XML (`tests\ribbon-check.ps1` checks `onLoad` and screenshots the tab). On narrow windows with many add-in tabs PowerPoint hides the *rethink* tab label at the right end of the tab row – widen the window.

## Gantt in one page

Behaviour follows the public user manual of the original (chapter "Gantt chart"); nothing else was used.

- **Datasheet** (`table`): row 0 = column headings (Aktivität, Start, Ende, Meilenstein, Verantwortlich, Bemerkung, Kategorie, Vorgänger, Form, Marker – the headings of the label columns appear in the chart), row 1 = milestone row below the chart, from row 2 one activity per row. *Form* (Balken, Abgerundet, Fünfeck, Chevron) and *Marker* (Raute, Dreieck, Kreis, Quadrat, Stern) override the design per row. A date without time is 00:00, so an end date of the 20th does **not** include the 20th (as documented for the original). Cell colours colour the bar (start cell) or milestone (milestone cell).
- **Free items** (`chart.items`): `bar`, `pentagon`, `chevron` (with `text`), `milestone`, `bracket`, `shade` – each with `row`, `start`, `end`, optional `text`, `fill`. Rows are `table` indices; inserting/deleting/moving rows through rethink moves the items along.
- **Scales:** year, quarter, month, week, day; automatic (finest unit whose labels fit + its parent) or fixed (`scales`), formats per scale (`formats`) using the documented date codes (`d dd ddd dddd ddddd w ww www wwww M MM MMM MMMM MMMMM Q QQ QQQ yy yyyy`, `"text"`, `\x`, `'`, `¶`). Separators: solid for the coarsest shown scale, dashed for finer ones. Weekends plain/shaded/hidden (hidden = no width on the axis), weekend days and week start configurable. Default range of a new chart: current month + the next two; without a stored range it follows the data (whole months).
- **Layout:** label columns (Activity always, Responsible/Remark when they have entries or are switched on), bold headlines at the height of the finest scale, bars with arrows where they leave the range, milestones (diamond) mid-day when the day scale is shown, otherwise at the start of the day, labels placed automatically without overlapping anything in the row, milestone row with its own rule, brackets at the top of a row, shades behind everything with their caption below the chart, optional row shades, row separators, milestone lines, today line (`today`, rethink extension).
- **Add-in interaction:** select the chart → blue glow with grips at both ends of every bar/bracket/shade and invisible move areas on bars and milestones; dragging snaps to days (to weeks/months when days are narrower than ~3 px) and shows a tooltip *05.11.2026 (Mi der KW 45)*; the visible range stays fixed while dragging; one undo entry per drag; Esc cancels. Click on a bar/milestone selects it inside the group (for formatting). Arrows at both scale ends drag the range in units of the finest scale (applied on release). Double-click on a scale → calendar for the range; on an item → its calendar (milestones close after one pick); elsewhere → datasheet. Right-click → *rethink-Zeitplan*: insert/duplicate/delete/move row, shade row, new bar/milestone/bracket/shade/pentagon/chevron at the mouse position (into the datasheet when the row has none yet), delete item, calendar, fit scale to data, scales (year … day, automatic), display toggles (weekends shaded/hidden, separators, dates on bars/milestones, milestone lines, today line, row separators, headlines, Responsible/Remark columns), datasheet.
- **Right-click (since 2026-10-19):** every chart has an invisible, clickable background shape (`background`, fill 100 % transparent) – without it a right-click on an empty spot hit no shape, PowerPoint deselected the chart and showed the slide menu. A right-click on a Gantt (selected or not) is caught by the mouse hook and opens rethink's own menu (`addin/gantt_menu.cpp`, built from `GanttState`): form of the clicked bar/milestone (Spalte 8/9 of its row, or the item kind), *Verbinden mit* (predecessors), row commands, new items at the mouse position, scales on/off and **format per scale** (examples with today's date), separators vertical/horizontal, milestone lines, today line, weekends normal/light grey/hidden, date labels and date format, columns (Kategorie, Verantwortlich, Bemerkung), bar height and default forms, category style/colours/sorting, connections, design, calendar, fit, datasheet. PowerPoint's own context menus keep the *rethink-Zeitplan* submenu as a fallback (e.g. Shift+F10).
- **Designs** (`design`: standard, modern, contrast, minimal): templates for bar height/shape, marker, category style, header band, row stripes, separators; all colours are theme colours (accents, tx1/bg1/tx2 with brightness), so charts follow the slide master. Single options (`barHeight`, `barShape`, `milestoneShape`, `separators`, `rowSeparators`, `categoryStyle`, `categoryColors`) override the design; choosing a design resets them.
- **Categories:** datasheet column 7 (*Kategorie*, index 6) groups consecutive rows: bold label with bracket (Standard), coloured block (Modern/Contrast) or text, a line between groups, optional bars in category colours (milestones a darker shade). Datasheet context menu and chart menu: *Nach Kategorie sortieren …* (dialog to choose the order; stable; free items and row shades move with their rows; `sortCategories=A|B` for automation).
- **Standard design = the documented look of the original:** thin bars (30 % of the row), ▲ milestones with the date centred below, bold left-aligned scale labels without ticks, dotted week lines and solid month lines, frame lines left and right, dashed milestone lines from the rows down to the marker; weekends plain by default. Colours always come from the slide master's theme, one theme slot per element kind: bars `accent1`, milestones and their lines `accent2`, process arrows (pentagon/chevron items) `accent3`, brackets `accent4`, shades light `accent1`, connections `tx2`, today line `accent6`; grid and text stay `tx1`. The mini toolbar shows that default as the swatch of *Automatisch (Master)*.
- **Per-row and per-item look** (`rowStyles: [{row, bar:{fill,line,weight,dashed}, milestone:{…}}]`, item `style`): colours are theme names (`tx1`, `bg1`, `tx2`, `bg2`, `accent1…6`), `#rrggbb` or `none`. Kept through datasheet edits, row moves, sorting and datasheet row insert/delete. The in-place update only rewrites fill/outline when the model's style changed (`RETHINK_STYLE` signature), so manual PowerPoint formatting survives.
- **Mini toolbar** (`gantt_menu.cpp`, `showBar`): appears above the right-click menu, like the original. Bar: *Balken / Gestrichelter Balken* + forms, fill and outline colour (swatches read from the slide's `ThemeColorScheme`, *Weitere Farben …* opens the colour dialog), line weight and dash. Milestone: marker, fill, outline, date on/off. Date label: content (none/range/start/end) and date format with today's date as example. Scale: scales on/off and format per scale. Commands `fill=`, `line=`, `lineWeight=`, `dashed=`, `barStyle=solid|dashed` act on the clicked item. A dashed bar keeps its colour in the outline and gets it back when switched to solid.
- **Radial menu** (`radial_menu.cpp`): right-click on an empty spot of the timeline opens round buttons around the cursor (bar, process arrow, milestone, bracket, shade, *Weitere Optionen …* for the full menu, × in the middle) with red crosshairs over the timeline at the insert position. Everywhere else the right-click shows mini toolbar + menu.
- **Drag rows:** invisible handles over the activity labels; dragging a row up/down shows a blue insertion line and moves the row on release (`moveRow=N`, `moveGanttRowTo`). Dropping into another category sets that category in the datasheet; the built-in datasheet and Excel follow. Inserting/deleting rows in the built-in datasheet shifts free items, row shades and row styles (undo shifts back; undoing a delete brings the rows back but not their row styles).
- **Row grip and cells (since 2026-10-19):** a row is moved at the faint six-dot grip left of the chart (it lights up, together with a very light band over the row's labels, when the mouse is over the row; tooltip „Zeile verschieben“ / „→ nach Planung“ while dragging); clicking a label no longer selects an inner shape, so no resize circles. Every label cell has its own overlay handle (`CellEdit`, I-beam cursor, hover glow = light fill + 1 px frame). A click opens an in-place editor (`events.cpp`, borderless EDIT popup over the cell, Segoe UI at font size × zoom, all selected): Enter commits, Tab/Shift+Tab commit and move to the next/previous label column of the row, Esc cancels, losing focus commits; commit = command `cell=<row>|<col>|<text>` (`ganttSetCell`, split at the first two `|`). An empty column that was just switched on gets at least `6 × line height` width.
- **Individual separator lines:** `rowLines` (table rows with a line at their top edge; ≥ 3, never above the first activity row; the line belongs to its row and moves with it through insert/delete/move/sort, a line that would end up above the first row is dropped) and `dateLines` (ISO days, line at the start of the day), drawn in the grid colour, 0.75 pt solid, roles `line row rN` / `line date YYYY-MM-DD`. Toggle commands `rowLine=N`, `dateLine=YYYY-MM-DD`. The selected chart highlights the boundary within 4 px of the mouse (between activity rows; edges of the finest shown scale inside the plot) with a light blue band; positions are stored in the `RETHINK_GANTT` tag (columns, bottom, scale edges), no layout per mouse move. A right-click within ~3 px of a boundary skips the radial menu and puts „Waagerechte/Senkrechte Trennlinie hier einfügen/entfernen“ at the top of the menu. Not done: a line below the last row (the milestone row rule is there), hover highlight while the chart is not selected.
- **Calendar** (`calendar_window.cpp`): year grid, one month per row aligned by weekday (Monday first), grey weekend columns, today in red, drag over days to pick a range, grab the start or end day to move it, Shift+click moves the nearer end; summary *01.06.2026 bis 30.06.2026 · 4 Wochen 2 Tage* and the hovered day with its ISO week; month back/forward, month dropdown, year spinner, mouse wheel.
- **Scale selected** (second click on a scale label): blue band over the scale with ◀▶ handles at both ends for dragging the range.
- **Connections:** column *Vorgänger* (index 7, activity names separated by `;`) draws arrows from the predecessor's end to the successor's start; *Verbinden mit* in the menu edits that cell.
- **Code:** `src/engine/dates.*` (calendar maths, formats), `src/engine/gantt.cpp` (layout, `GanttAxis`), `src/engine/gantt_edit.*` (role → cells/items, drag/snap, range, rows, items – pure and unit-tested), `src/ppttc.cpp` (`parseGantt`, `ganttEditsJson`), `addin/elements.cpp` section "Zeitplan" (modify/drag/hit test/commands/calendar), `addin/events.cpp` (handles, gestures, double-click, cursor hit), `addin/calendar_window.*`, `addin/gallery.*`, context menu generated in `addin/addin.cpp` (`ganttMenuXml`, inserted at the `<!--rethink:gantt-menu x-->` markers of `ribbon.xml`; ids get a suffix per PowerPoint context menu because ribbon ids must be unique).

### Gantt – not yet like the original (next candidates)

- Anchors between items (handle dropped on another item's anchor point), Ctrl+drag duplicate, Shift+Ctrl+drag duplicate-and-anchor.
- Dragging items vertically into another row; keyboard shortcuts (Del, F2, Ctrl+Alt+↑↓, Shift+Alt+↑↓/←→).
- Red crosshairs with date tooltip while merely hovering (today only while the radial menu is open).
- Scale labels ("Add Scale Label"), fiscal years, label content per item (today one setting for all bars), manual label positions, tints of theme colours in the colour menu.
- Mini toolbar, radial menu and year calendar are tested through automation (`GanttToolbar`, `GanttRightClick`, `CalendarPick`); their look still needs a check by hand with a real mouse (the test screen was locked).
- Label columns: multi-row Responsible/Remark labels with brackets, indentation hierarchy, checkboxes and Harvey balls, pasting multi-line text into a column.
- Right-to-left layout; Excel range link (rethink offers the Excel datasheet instead).

## Build & run

```bat
:: Developer Command Prompt (vcvarsall.bat x64), in the repository folder
cmake -S . -B %TEMP%\rtk -G "NMake Makefiles" -DCMAKE_BUILD_TYPE=Release
cmake --build %TEMP%\rtk
```
- Copy `rethink.exe` and `rethink-addin.dll` into `bin\` (create it; it is git-ignored). **Close PowerPoint first** – it locks the registered DLL.
- `scripts\register.ps1` registers the add-in for the current user (HKCU only); `scripts\unregister.ps1` removes it.
- Use a short build directory: MSBuild fails on paths > 260 characters.
- On ARM64 Windows build the **x64** target (`vcvarsall.bat x64`): Office runs as x64 under emulation and can only load x64 add-ins.

## Tests

| Command | Covers |
|---|---|
| `%TEMP%\rtk\engine_test.exe` | engine: waterfall arithmetic and orientations, layout bounds, dates/weeks/formats, Gantt layout, Gantt editing (drag/snap/range/rows/items), datasheet XLSX round trip, cell input |
| `%TEMP%\rtk\addin_selftest.exe %TEMP%\rtk\rethink-addin.dll` | COM class factory, ribbon XML incl. generated context menus and unique ids, callbacks, no PowerPoint needed |
| `tests\run-tests.ps1 -Exe bin\rethink.exe` | CLI validation and error messages (incl. Gantt options and items) |
| `bin\rethink.exe build tests\gantt.ppttc -o tests\out\gantt.pptx`, then `tests\export-png.ps1 tests\out\gantt.pptx` | Gantt and horizontal waterfall from `.ppttc`, PNGs for a visual check |
| `tests\waterfall-addin-test.ps1` | waterfall in PowerPoint (insert, handles, drag, undo, datasheets, Excel, resize, orientation, build-down) |
| `tests\gantt-addin-test.ps1` | Gantt in PowerPoint (insert, overlay handles, drag end/move/milestone/range, Esc, undo, datasheet sync incl. row insert/delete, context commands and toggles, row drag into another category, mini toolbar and styles, radial menu, calendar pick, resize) |
| `tests\gantt-mouse-test.ps1` | real mouse input on the overlay (takes over the mouse for a few seconds) |
| `tests\ribbon-check.ps1` | PowerPoint accepts the ribbon XML; screenshot of the tab |
| `tests\addin-test.ps1 -OutDir %TEMP%\rtk-addin-test` | older add-in commands; keep `-OutDir` outside OneDrive (PowerPoint reports OneDrive files as https URLs) |

Run the PowerShell tests in their own process (`powershell -NoProfile -ExecutionPolicy Bypass -File …`). The scripts are UTF-8 **with BOM** so Windows PowerShell 5.1 reads the umlauts; keep it that way for new scripts.

Automation cannot click; every interaction has an entry point on `COMAddIns("rethink.addin").Object` that runs the same code path: `InsertElement(kind, orientation)`, `AdjustSegment`, `DragSegment`, `GanttDrag(role, grip, "px;…", cancel)`, `GanttCommand(id, x, y)`, `GanttToggle(id, pressed, x, y)`, `GanttPressed`, `GanttHit(x, y)`, `GanttHover(x, y)`, `GanttEditCell(x, y, text)`, `GanttTargets`, `SetGanttDates(role, first, last)`, `OpenCalendar`/`CloseCalendar`, `OverlayInfo`, `OpenDatasheet`, `DatasheetEnter`, `DatasheetPaste`, `DatasheetCell`, `CloseDatasheet`, `OpenExcelDatasheet`, `DatasheetWorkbook`, `ApplyDatasheets`, `EventStatus`, `RibbonStatus`, `ActivateTab`. Add one for each new interaction.

## Architecture in one page

```
.ppttc JSON ──► src/ppttc.*  ──► DataItem {name, table, chartJson}
                                     │
                       src/engine/*  ▼  layoutChart(item, options, box) → Drawing {prims[], axis, rows}
                                     │                     (pure C++, unit-tested)
              ┌──────────────────────┴───────────────────────┐
   src/render_element.* (OOXML group + RETHINK_MODEL tag)   addin/elements.* (PowerPoint object model,
   used by the CLI (src/build.cpp)                           role-named shapes, in-place update, tags)
                                                                   │
                       addin/events.* (selection/double-click events, overlay tracking, handle gestures)
                       addin/overlay.* (layered window: glow, label, handles, tooltip, table icon)
                       addin/datasheet_window.* (built-in spreadsheet-like datasheet)
                       addin/calendar_window.* (range/date calendar), addin/gallery.* (insert menu)
                       addin/addin.cpp (COM add-in, ribbon + context menu callbacks, automation interface)
```

Key rules:
- The engine never touches COM or XML; renderers never compute layout. Edits of the model (Gantt drag, rows, items) are engine functions too (`gantt_edit.h`).
- Shapes inside a rethink chart are named by role (`segment r1 c3`, `bar r2`, `milestone r1`, `item 0`, `scale month 2`, …). Same roles ⇒ update in place; different roles ⇒ full redraw (keeping name, z-order and the selection).
- Group tags: `RETHINK_MODEL` (JSON data item), `RETHINK_LAYOUT` (layout box + group rect when last drawn), `RETHINK_KIND` (`waterfall up|down|right|left` or `gantt`). Child tags: `RETHINK_STYLE` (last colour set by rethink, to detect user formatting), `RETHINK_SEG` (`pos`/`neg`/`e`; Gantt: `bar`/`milestone`/`bracket`/`shade`).
- During a handle drag the overlay holds the mouse capture: never show a message box mid-drag (collect the error, report after `DragEnd`), and the prim cache (`samePrim` against the last drawing) is only valid for that reason.
- WinRT JSON, OPC and MSXML run on a worker thread (`onWorker`) – never initialise WinRT on PowerPoint's UI thread. Pure layout may run on the UI thread (drag steps do).
- Never redraw or select inside a mouse handler of the overlay; post a message and act afterwards. Never block on Excel; treat `RPC_E_CALL_REJECTED` / `RPC_E_SERVERCALL_RETRYLATER` as "retry next tick".
- Layered overlay pixels with alpha 0 are click-through even if `WM_NCHITTEST` says otherwise: invisible handles are drawn with alpha ≈ 1/255.

## Conventions

- C++20, MSVC `/W4 /permissive- /utf-8`, the add-in builds with `/WX`.
- UI strings are German (the current user's language); repository documents and commit messages are English.
- Clean-room: use the documented behaviour only; no decompiling, no think-cell names, identifiers, icons or artwork; reverse-engineering notes stay out of the repository.
