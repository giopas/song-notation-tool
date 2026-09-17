# Changelog

All notable changes to this project are documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.0.0/).

## [0.18.1] — 2026-09-17

A follow-up to `webserver.py`'s launch mode from v0.18.0, closer to how
the sibling `qlc-plus-swiss-knife-tool-script` project's window
actually behaves: no OS titlebar by default, and a proper in-app way to
quit now that there's no titlebar close button to rely on.

### Added
- **Frameless native window by default.** When `pywebview` is
  installed, the window it opens now has no OS titlebar at all
  (`frameless=True`), not just no browser chrome. `easy_drag=True`
  keeps it movable by clicking anywhere in it; `confirm_close=True`
  keeps Cmd/Ctrl+Q as a safety-netted fallback.
- **"⏻ Save & Close" button** in the front end's toolbar — saves the
  currently open song, then calls the new `POST /api/quit` route to
  shut the app down cleanly. This is now the normal way to quit,
  since a frameless window has no close button of its own.
- **`POST /api/quit`** — destroys the native window in webview mode;
  in browser mode, stops the HTTP server via `socketserver`'s own
  thread-safe `shutdown()`.
- **Clearer terminal output when `pywebview` isn't installed.** Instead
  of silently opening a browser tab, `webserver.py` now prints the
  install command (`pip install pywebview`, or `pip install -r
  requirements-optional.txt`) and reminds you of `--browser` /
  `--no-open`. Only shown when the fallback is actually due to the
  missing package — not when `--browser` was passed on purpose, and
  not with `--no-open` (nothing was going to open either way).

### Fixed
- The first pass at `/api/quit` (written but never released) used a
  self-sent `SIGINT` to stop the server in browser mode, matching the
  reference project. Caught during testing: self-directed signal
  delivery to a backgrounded process isn't reliable in every
  environment — confirmed with an isolated repro where even an
  externally sent `SIGTERM` didn't stop a trivial Python loop, only
  `SIGKILL` did. Switched to `socketserver.BaseServer.shutdown()`,
  which needs no signal delivery at all and works the same everywhere.

## [0.18.0] — 2026-09-17

Two things: the app now runs headless and in a browser, not just as a
Tkinter window; and the Tkinter app itself is more self-explanatory to
a new or returning user. See the UX & Multi-Platform Enhancement Spec
for the full brief this closes, and `ROADMAP.md`'s former "§11's
architecture fork — Tkinter vs. Flask + browser SPA" item, which this
release resolves.

### Added

- **`cli.py`** — a headless CLI with no window at all, built entirely
  on the pure-data layers:
  - `cli.py convert -i song.sng -e {txt,pdf} [-o out] [--instrument …] [--orient …]`
    — convert one `.sng` file.
  - `cli.py batch -i songs/ -e {txt,pdf} [--out-dir …]` — re-export
    every `.sng` in a folder, e.g. before a gig or after a formatting
    change; reports per-file success/failure and exits non-zero if any
    failed.
  - `cli.py lint -i song.sng` — parse every section's chart line and
    report errors (or confirm they all parse cleanly) without opening
    the GUI.
- **`webserver.py`** + **`web/`** — a thin, stdlib-only local web
  server (`http.server`, zero external dependencies, matching the
  project's own rule) exposing the core engine over HTTP/JSON, plus a
  static HTML/CSS/JS front end: song list, meta form, a chart-line
  editor per section with live-parsing preview (mirrors the desktop
  chart editor bar), reorder/duplicate/delete, a collapsible "how this
  works" help strip, a first-launch "Start here" panel, and TXT/PDF
  export. Run `./webserver.py` (serves `./songs` on
  `http://localhost:8420`) or `./webserver.py --host 0.0.0.0` to reach
  it from another device on the network (e.g. a phone at rehearsal).
  API routes are documented inline in `webserver.py`.
  - Running it now opens a window automatically instead of leaving you
    to copy the URL into a browser by hand — a native, chrome-less
    window via the optional `pywebview` package (`pip install
    pywebview`; matches the sibling
    [qlc-plus-swiss-knife-tool-script](https://github.com/giopas/qlc-plus-swiss-knife-tool-script)
    project's launch UX) when it's installed, falling straight back to
    your default browser tab when it isn't — nothing about this is a
    *required* dependency. `--browser` forces a normal browser tab even
    with `pywebview` installed; `--no-open` starts the server without
    opening anything (e.g. when only serving other devices on the
    network).
- **`export.py`** — `build_song_lines(doc, instruments=None)` and
  `build_pdf(doc, instruments=None, orient="portrait")`, extracted
  from `SongNotationApp._build_song_lines`/`_build_pdf`/`_assemble_pdf`.
  Pure functions over a `doc` dict (no Tkinter import anywhere in the
  module — covered by a dedicated test), so the CLI and web server call
  the exact same TXT/PDF builders the desktop app uses; a fix or a new
  export field now only has to be made once.
- **`constants.py`** — `APP_VERSION`, `INSTRUMENT_STRINGS`,
  `SECTION_TYPES`, `RENDER_MODE_LABELS`, `TAB_BEATS_DEFAULT`,
  `TAB_BEATS_OPTIONS`, and `default_export_name()` moved out of
  `song_writer.py` so non-UI code (CLI, web server, tests) doesn't need
  to import Tkinter just to read an instrument's string list.
- **`examples.py`** — the built-in "Example Song" sample behind "Open
  example," shared by the desktop app and the web server so it can
  never drift between the two.
- **Desktop app — "Start here" panel.** First launch (or a brand-new
  document with no sections) shows *New song*, *Open example*, and
  *Import project* instead of a blank grid.
- **Desktop app — help strip.** A "?" button in the toolbar opens a
  plain-language explanation of sections, chart-line syntax, repeats,
  riffs, duplicate-as-reference, and transpose.
- **Desktop app — live Preview window.** A "👁 Preview" toolbar button
  opens a window showing the TXT/PDF export as it will look, refreshed
  automatically on every edit that touches the song map.
- `tests/test_export.py` — 7 new tests covering `export.py`, including
  an explicit assertion that the module has no `tkinter` import.

### Changed

- `song_writer.py`'s `_build_song_lines`/`_build_pdf` are now thin
  delegators to `export.py` (`_sync_doc_meta()` copies the live
  StringVars into `self.doc["meta"]` first, same as `_save()` already
  did) — the duplicate TXT/PDF-building code that used to live in the
  Tkinter class is gone.
- `_save()` and the export dialogs use `constants.default_export_name()`
  instead of each re-computing the "`<Artist> - <Title>`" filename
  stem inline.
- `APP_VERSION` bumped to `0.18`, sourced from `constants.py` everywhere
  (desktop app, CLI, web server) instead of being defined separately.

## [0.17.0] — 2026-09-16

The UI rewrite `DESIGN_v0_17.md` called for: the app now edits the
v0.16 document model directly — a section's `items` — instead of the
retired per-measure `layers` grid. The song map, the chart editor bar,
and the riff library become the editing surface for every
`render:"chart"` section (most of them); the old measure grid survives,
rewired onto `measure` items, for `render:"tab"` sections. 92 unit
tests, all independent of Tk/a display.

### Added
- **Song map** — the section listbox + single-section editor split is
  replaced by a scrollable list of rendered section rows (fret line over
  symbol line), the whole song visible at once. Click a row to focus it.
- **Chart editor bar** — one live-parsing text field is now the entire
  editing surface for `render:"chart"` sections: type, it re-renders on
  a ~150ms debounce, a parse error leaves the last valid render on
  screen and shows inline instead of clearing anything. Tab commits and
  advances to the next section; Escape reverts; Enter commits and adds
  a new section below.
- **Riff library strip** — create/rename/delete riffs from the UI for
  the first time (the `block_ref` grammar already existed; there was no
  editor for `blocks`). Editing a riff updates every referencing section
  because there is only one copy of the data. Deleting a referenced
  riff is blocked and names the referencing sections.
- **Promote to riff (Ctrl/Cmd+R)** — select a run in the chart editor
  bar and turn it into a named riff in place.
- **Duplicate as reference (Ctrl/Cmd+D)** — inserts a `section_ref`
  rather than a copy, removing the duplicate-page failure mode the old
  copy dialog produced.
- **Reordering** — drag a row by its name, or Alt+↑ / Alt+↓. No dialog.
- **Stage View (Ctrl/Cmd+P)** — full-window, high-contrast, read-only,
  built from the same row renderer as TXT/PDF.
- **Live page-count indicator** in the header, and a page-break-before-
  section rule in the PDF builder so a section is never split across a
  page.
- `songmap.py` — pure, unit-tested riff/reorder/promote/duplicate logic.
- `render.py`: `resolve_display_items()`, `mark_spans()`, `has_coda()`,
  `estimate_page_count()` / `estimate_section_lines()`.

### Changed
- The UI now reads/writes the v0.16 document model (`self.doc`, shaped
  like `model.py`'s schema) directly. The v0.15-era `Section` class and
  its `layers` dict are retired; `.sng` files still round-trip through
  `model.migrate_document()` for anything saved by an older version.
- **Transpose is non-destructive** — the Transpose dialog sets a
  `document`/`section` offset instead of rewriting stored fret numbers;
  every renderer resolves it at display time. A `+n` then `-n` round
  trip is exact by construction.
- Export dialogs dropped the "include layers" checkboxes (chords/notes/
  lyrics no longer exist as separate layers — a chart line already
  carries chord/note symbols); "include instruments" stays.
- Marks (`|:`, `:|`, 1st/2nd endings, segno, coda, D.C./D.S., `simile`)
  now actually render, inline, wherever they sit in a chart line — in
  the map, TXT, and PDF alike.

### Removed
- The v0.15 section-link dialog (`link_id`) — superseded by
  `=sectionname` references typed directly in the chart line, and by
  Ctrl/Cmd+D.
- The custom-drawn scrollbar indicator canvases (v0.6) — simplified to
  standard `ttk.Scrollbar`s in the rewrite; functionally equivalent,
  the drawn-canvas look is gone.

### Deferred
Being upfront about the gap between `DESIGN_v0_17.md` and this
release — graphical 1st/2nd-ending brackets and a boxed coda block
(`render.mark_spans()`/`has_coda()` resolve the data; nothing draws
the graphics yet), chart-editor-bar autocomplete, and tab-grid
measure copy/paste all stayed out of scope here. Tracked live in
`ROADMAP.md` rather than repeated per-release.

## [0.16.0] — 2026-09-16

v0.15 supported exactly one thing — a full measure-by-measure tab grid
— which meant every song came out as several pages, mostly dashes,
with repeated sections written out in full (see `DESIGN_v0_16.md`
section 2 for the audit against a real song). This release is the
foundation for the compact, one-page charts the project is actually
for.

### Added
- New `.sng` data model (format 2): a section is a flat sequence of items
  (`token`, `group`, `block_ref`, `section_ref`, `measure`, `mark`), covering
  the full vocabulary of the hand-written reference sheets. See
  `DESIGN_v0_16.md`.
- **Chart grammar** — one text field per section (`grammar.py`) parses a line
  like `5A 5D [5A 5D]x2 riff1 x3 +2` into structured items, with a matching
  unparser so the saved file stays the single source of truth.
- **Render-time transposition** (`transpose.py`) — `document.transpose +
  section.transpose + ref.transpose`, applied only when drawing/exporting.
  Never rewrites stored notes; a `+n` then `-n` round trip is exact.
- **v0.15 → v0.16 migration** (`model.migrate_document`) — old files load
  without loss; `link_id` becomes a `section_ref`, tab/chords/notes layers
  become items. Nothing destructive: a v0.15 `.sng` opens as-is and the
  app won't overwrite it with the new format until you explicitly save.
- A "Chart line" field in the section editor, parsed live against the new
  grammar (minimal UI hook — the full chart-row editor and song map are
  v0.17).
- `examples/` with a migrated `sample_song.sng` and the hand-written
  reference `sample_song.txt` / `sample_song.pdf` this project is trying to
  reproduce electronically.
- `CHANGELOG.md`, `DEVELOPMENT.md`, this repo's first tagged release.

### Fixed
- `ENHARMONIC` had `"Bb"` defined twice and was missing the lowercase flat
  variants (`db`, `eb`, `gb`, `ab`). The chromatic table now lives in one
  place (`transpose.py`) and both exporters import it.
- TXT and PDF export disagreed on whether a section "has tab" — each
  computed it inline as `any(list_of_dicts)`, which is `True` for almost
  any section since a non-empty dict is truthy regardless of its values.
  Both now call the same `render.section_has_tab()`.
- TXT export defined `W = 80` and never used it — every section was emitted
  as one long unwrapped line. `W` (default 100) now actually drives
  line-wrapping, matching the measures-per-line logic the PDF builder
  already used.
- Fret entry accepted leading zeros (`01`) and values above `MAX_FRET`;
  both are now rejected in the tab grid and in the chart grammar.
- PDF export defaulted to landscape; compact mode targets one page, so the
  export dialog now defaults to Portrait A4.

### Changed
- `song_writer_v.0.15.py` renamed to `song_writer.py`. The version lives in
  `APP_VERSION` and on the release asset — the version-in-filename broke
  every link and script on each release.
- Compact export mode (default on): strings with no content anywhere in a
  section are dropped from TXT and PDF output.

## [0.15] — ongoing refinements before this changelog existed

## [0.14]
- Minor stability and layout fixes.

## [0.13]
- Beats selector moved into the editor toolbar (always visible).
- Hamburger ☰ menu for narrow windows.

## [0.12]
- Variable tab beats (8/16/32); protected dashes; smart transposition.

## [0.11]
- Chromatic note list with enharmonic aliases; transpose all layers;
  root-note picker.

## [0.10]
- TXT/PDF export fixes; Artist before Title; cleaned-up chord/note display.

## [0.8]
- Resizable left panel sash; measure grid wrap mode.

## [0.7]
- Section linking — propagate edits across linked sections.

## [0.6]
- Fixed macOS layer-toggle colours; custom scroll indicators; copy/paste fix.

## [0.5]
- Fixed macOS button colours; horizontal tab Entry per string with dashes.

## [0.4]
- Two-row topbar, Drop D toggle, Transpose dialog.

## [0.3]
- Artist name field, section copy dialog, APP_VERSION constant.

## [0.2]
- Light/dark theme engine, ToolTips, ttk styling, layer toggle buttons.

## [0.1]
- Initial release — sections, layers (tab/chords/notes/lyrics), export TXT.
