# Song Notation Tool v0.17.0

The UI rewrite `DESIGN_v0_17.md` called for. The app now edits the v0.16
document model directly — a section's `items` — instead of the retired
per-measure `layers` grid. The song map, the chart editor bar, and the
riff library are the new editing surface for every `render:"chart"`
section, which is most of them; the old measure grid survives, rewired
onto `measure` items, for `render:"tab"` sections.

## Highlights

- **Song map** — every section on screen at once as a rendered chart row
  (fret line over symbol line), not one section at a time behind a
  listbox. Click a row to load it into the editor bar; the repeat count,
  any annotation, and a coda marker render inline on the row itself.
- **Chart editor bar** — one text field, live-parsing on every keystroke
  (~150ms debounce). A parse error never clears the row: the last valid
  render stays on screen and the error shows inline underneath. Tab
  commits and moves to the next section; Escape reverts; Enter commits
  and adds a new section below.
- **Riff library strip** — create a riff, reference it from any section
  with its name, and edit it once to update every section that
  references it (there is only ever one copy of the data — nothing to
  keep in sync). Deleting a riff that's still referenced is blocked and
  names the sections holding the reference.
- **Promote to riff (Ctrl/Cmd+R)** — select a run of tokens in the chart
  editor bar and turn it into a named riff in place, the way a riff
  actually gets created in practice: written out once, then noticed to
  repeat.
- **Duplicate as reference (Ctrl/Cmd+D)** — inserts `=sectionname`
  rather than a byte-for-byte copy. This is what removes the duplicate
  Verse 2 / Chorus 2 pages the v0.15 copy dialog produced.
- **Reordering** — drag a row by its name, or Alt+↑ / Alt+↓ with no
  mouse at all. No dialog either way.
- **Non-destructive transpose** — the Transpose dialog now sets a
  `document`/`section` offset instead of rewriting stored fret numbers;
  every renderer (map, TXT, PDF, Stage View) resolves it at display
  time via `transpose.resolve()`. A `+5` then `-5` round trip is exact,
  by construction, because nothing was ever rewritten.
- **Live page-count indicator** — `render.estimate_page_count()` gives a
  running "[N pages]" estimate in the header as you type, using the same
  line-budget geometry the PDF builder uses.
- **No section splits across a page break** — the PDF builder now checks
  a section's estimated height before starting it and breaks the page
  first if it won't fit, rather than mid-section.
- **Stage View (Ctrl/Cmd+P)** — full-window, high-contrast, read-only,
  built from the exact same `render.render_chart_row()` /
  `resolve_display_items()` calls the TXT and PDF exporters use — one
  renderer, three outputs, not three implementations.
- **Marks render** — barlines, 1st/2nd endings, segno, coda, D.C./D.S.,
  and `simile` all appear inline wherever they sit in the chart line,
  in the map, TXT, and PDF alike, because they're ordinary items in the
  same sequence `render_chart_row()` already draws.

## Deferred from the design doc

Being upfront about the gap between the design doc and this release:

- **Graphical ending brackets / boxed coda.** `render.mark_spans()` and
  `render.has_coda()` resolve the data a canvas or print-stylesheet
  renderer would need to draw an actual bracket over a 1st/2nd ending or
  box a coda block. This release renders marks as inline text tokens
  (`|1.`, `|2.`, `coda`, …) rather than drawing the graphics — the data
  is there for a later pass to draw from.
- **Autocomplete in the editor bar** (`=` for section names, bare
  identifiers for block names) — not implemented. Typos in a reference
  surface as an ordinary parse error instead.
- **Copy/paste of a single measure and a per-measure beats picker** in
  the tab grid — the v0.15 editor had both; this rewrite keeps the grid
  functional (add/remove a bar, edit any cell) but drops those two
  conveniences for now.
- **The custom-drawn scrollbar indicator canvases** (added in v0.6 for
  macOS Aqua theming) were simplified to standard `ttk.Scrollbar`s in
  the rewrite. Functionally equivalent; the drawn-canvas look is gone.
- **§11's architecture fork** (Tkinter vs. Flask + browser SPA) — this
  release stays in Tkinter. The `model`/`grammar`/`transpose`/`render`
  layers have no Tk dependency either way, so the fork is still open for
  a later release if the graphical-mark and drag-and-drop work above
  turns out easier in HTML/CSS, as the design doc suspected it would.
- **Empirical timing / one-page guarantee.** The one-page A4 target and
  the "under five minutes, keyboard-only" entry time are architectural
  goals this release is built toward (page estimate + page-break-before-
  section logic; Enter/Tab-driven entry), not measurements taken against
  a real song.

## Under the hood

- `songmap.py` (new) — pure, unit-tested logic for riff usage tracking,
  blocked deletion, promote-to-riff, section reordering, and
  duplicate-as-reference. No Tk dependency, same as `model`/`grammar`/
  `transpose`/`render`.
- `render.py` gained `resolve_display_items()` (the shared render-time
  transpose walk), `mark_spans()`, `has_coda()`, and the page-count
  estimator.
- 92 unit tests (`tests/`), all independent of Tk/a display.
