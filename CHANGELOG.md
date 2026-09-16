# Changelog

All notable changes to this project are documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.0.0/).

## [0.17.0] — 2026-09-16

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
- The custom-drawn scrollbar indicator canvases (v0.6) — replaced with
  standard `ttk.Scrollbar`s.

See `RELEASE_NOTES_v0.17.0.md` for what's deferred from `DESIGN_v0_17.md`.

## [0.16.0] — 2026-09-16

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
  become items.
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
