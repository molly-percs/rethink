# rethink – charting ledger

This ledger lists the charting features rethink aims to cover, modelled on the **documented behaviour** of think-cell 14 (user manual chapters in brackets). The manual is used as a behavioural spec only; no text or images are copied.
Slide library and templates are intentionally **out of scope**.

**Status:** ✅ done and tested · 🟡 partial · ⬜ open
**Backend:** **D** = shapes drawn by rethink (layout engine) · **N** = native PowerPoint chart (only for `.ppttc` template filling)

> Full parity with think-cell is a large effort. rethink delivers it in vertical slices: each slice is one chart type or feature, done end to end (data → slide → add-in) and tested. While any row below is open, the scope is **not** complete.

## Architecture (decision)

- **One layout engine as the single source of truth** (`src/engine/`): pure C++20, no COM, no XML. It owns data semantics (`e`, `100%=`, percent/absolute, empty-row colour skip), scales, geometry, label texts, decorations and the data table, and produces a list of primitives (rect, line, text) in EMU with style roles.
- **Thin renderers** turn primitives into shapes: **OOXML** (CLI, `.ppttc`) as a group in slide XML, and **COM** (add-in) via the PowerPoint object model. Neither computes layout.
- **Model on the element:** the group carries the tag `RETHINK_MODEL` (JSON: type, options, table), so the add-in can redraw it and `.ppttc` can refill it. Each shape is named after its role, which enables in-place updates.
- **Why drawn instead of native:** Mekko with variable widths, Gantt, CAGR/difference arrows, totals above stacks, waterfall connectors and a data table grouped with the chart cannot be expressed with native charts; native tables cannot even be grouped.

## Data semantics

| Feature | Behaviour (spec) | Backend | Status | Test |
|---|---|---|---|---|
| Datasheet layout | row 0 = categories, column 0 = series names | D, N | 🟡 (N ✅) | engine_test, run-tests |
| Empty row | skips one colour of the fill scheme | D, N | 🟡 (N ✅) | engine_test |
| Absolute/percent | percent values converted via column sum (or `100%=`) and vice versa | D | ⬜ | |
| `100%=` row | optional totals row for percent calculation | D | ⬜ | |
| `e` in waterfall | computed segment; total down to the baseline; computed backwards in the first column; at most one `e` per column. **JSON convention:** `{"string":"e"}` | D | ✅ | engine_test |
| Number format | automatic 0–2 decimals, locale decimal separator, `decimals` option | D | 🟡 (no custom format patterns) | engine_test |
| Step size when changing a value | decimals of the value as entered (12 → 1, 12.5 → 0.1); `decimals` option wins | D | ✅ | engine_test |
| Date categories | year / month for 1 Jan / first of month | D, N | 🟡 (N ✅) | |

## Chart types

| Type | Key points | Backend | Status |
|---|---|---|---|
| Stacked column/bar | single series = simple column; bar = rotated; gap width | D (N ✅) | ⬜ |
| 100% stacked | sum may differ from 100%; percent and/or absolute labels; 100% marker | D (N 🟡) | ⬜ |
| Clustered | segments side by side | D (N ✅) | ⬜ |
| **Waterfall (build-up)** | `e`, several segments per column, connectors, data table | D | ✅ (connectors automatic only) |
| Waterfall (build-down) | same as build-up with descending default data (`e` in the first column) | D | ✅ |
| Chart orientation | rotate/flip via the insert menu (up, down, left, right) | D | ✅ waterfall (`orientation`) |
| Line / profile | interpolation, markers, value axis for increasing X, second Y axis | D (N ✅) | ⬜ |
| Area / 100% area | stacked, totals, interpolation | D (N ✅) | ⬜ |
| Combination | lines and columns per series | D | ⬜ |
| Mekko % axis | variable column width = column total, Y in %, no gaps | D | ⬜ |
| Mekko unit axis | "X extent" row for widths, absolute heights | D | ⬜ |
| Pie / doughnut | one category, rotation, inside/outside labels | D (N ✅) | ⬜ |
| Scatter / bubble | rows = points (X, Y, size, group), trend lines, partitions | D | ⬜ |
| **Gantt / timeline** | calendar scales, activities, bars, milestones, drag in days | D | ✅ core (see HANDOFF.md: anchors, crosshairs, scale labels, Harvey balls … open) |

## Decorations, labels, axes

| Feature | Behaviour | Backend | Status |
|---|---|---|---|
| Segment/series/category labels | automatic placement (inside, otherwise outside) | D | 🟡 segment + category for waterfall; no collision handling between outside labels |
| Totals above columns | total value above the stack | D | ⬜ |
| Waterfall connectors | horizontal line between adjacent columns | D | ✅ (re-hooking connectors manually ⬜) |
| Difference arrows | level and total difference; relative/absolute modes | D | ⬜ |
| CAGR arrows | (B/A)^(1/n) − 1, n from date categories (30/360) | D | ⬜ |
| Value lines | line at a value or mean, label left/right | D | ⬜ |
| 100% marker | for 100% charts | D | ⬜ |
| Series connectors | lines between segments of adjacent columns | D | ⬜ |
| Value axis | scale, ticks, gridlines | D | ⬜ |
| Category gap | widen individual gaps | D | ⬜ |
| Axis break | interrupt a value range | D | ⬜ |
| Legend | series, order | D | ⬜ |
| Segment order | sheet, reverse, ascending, descending | D | ⬜ |
| **Data table below the chart** | grouped, columns aligned with bars; rows only in the table or also in the chart | D | ✅ waterfall (`datatable`, `tableRows`) |

## Add-in interaction

| Feature | Status |
|---|---|
| Ribbon tab "rethink" | ✅ |
| Insert menu like the original (gallery of all chart types, orientation via arrows around each icon) | ✅ |
| Insert per chart type | 🟡 waterfall ✅, build-down ✅, Gantt ✅ |
| Built-in datasheet (spreadsheet look; table icon bottom right; double-click) with live update | ✅ typing, paste, undo, rows/columns, %, cell colour (automation-tested; manual mouse/keyboard check pending) |
| Datasheet in Excel (optional): live, closing never prompts | ✅ |
| Selection glow + tiny "rethink chart" label, document untouched | ✅ |
| ▲▼ handles per segment: click = one step (step follows the precision of the entered value) | ✅ |
| Drag handles to change values quickly (live, frozen scale, step by entered precision, `e` locked, Esc cancels, one undo entry) | ✅ (automation-tested in PowerPoint; ◀▶ when horizontal) |
| Glow follows the chart **while** it is being moved (hidden while resizing) | 🟡 implemented (mouse hook); manual check on Windows pending |
| In-place updates keep user formatting | ✅ |
| Resize with the mouse re-lays out the chart | ✅ |
| `.ppttc` fills rethink elements (`RETHINK_MODEL`) and placeholders | ✅ CLI |
| Gantt: drag bar ends/bars/milestones in days with snapping and date tooltip, drag scale ends, calendar (double-click), context menu (rows, new items, scales, display) | ✅ (automation + real mouse tested) |

## Slice plan

1. ✅ Waterfall + data table (engine, OOXML output, `.ppttc`, tests)
2. ✅ Add-in: insert, datasheet, overlay, handles, in-place updates
3. ✅ Drag handles, glow follows move (manual check pending), insert menu with all types + orientation, build-down waterfall
4. ✅ Gantt (calendar scale, bars, milestones, drag in days, datasheet) – remaining details in HANDOFF.md ← user testing
5. ⬜ Stacked, 100% and clustered columns incl. totals, value axis and legend
6. ⬜ Line, area, combination
7. ⬜ Mekko (% and units)
8. ⬜ Decorations: difference, CAGR, value line, 100% marker, series connectors
9. ⬜ Pie/doughnut, scatter/bubble
