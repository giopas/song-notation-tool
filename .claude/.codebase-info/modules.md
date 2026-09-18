# Modules

*Last Updated: 2026-09-18*

## Engine Modules (pure Python, no GUI dependency)

### model.py — Document Schema & Migration

The canonical definition of the `.sng` document format.

- `FORMAT_VERSION = 2` — current schema version
- `ITEM_KINDS` — valid item kind strings: token, group, block_ref, section_ref, measure, mark
- `MARK_NAMES` — valid mark names (barlines, repeats, endings, segno/coda/DC/DS, simile)
- Item constructors: `make_token`, `make_group`, `make_block_ref`, `make_section_ref`,
  `make_measure`, `make_mark`
- Document scaffolding: `new_document`, `new_section`, `new_block`
- `migrate_document(doc)` — auto-upgrades v1 → v2 (adds `blocks`, converts old section formats)
- `validate_block_name(name)` — identifier rules for riff names

### grammar.py — Chart-Line DSL

Parses the human-typed chart-line syntax into structured item lists and back.

- `parse(text) -> (items, annotation)` — full parse with optional annotation after `#`
- `parse_items(text) -> items` — items only
- `unparse(items) -> str` — reconstruct the text from items
- `ParseError` — raised on malformed input
- Token regex: `^(?P<fret>\d+)?(?P<note>[A-G])(?P<acc>[#b])?(?P<qual>[A-Za-z0-9/+\-()]*)$`
- Supports: tokens (`5A`, `G#m7`), groups (`[C D]x2`, `[E F]+3`), marks (`|:`, `:|`,
  `1.`, `2.`, `segno`, `coda`, `DC`, `DS`, `//`), block refs, section refs (`=chorus1`)

### render.py — Shared Rendering

- `render_chart_row(items, label) -> (fret_line, symbol_line)` — column-aligned two-line output
- `resolve_display_items(items, eff_semitones)` — apply transposition at display time
- `mark_spans(items)` — compute ending bracket spans for graphical rendering
- `estimate_page_count(doc)` — live page-count indicator
- `measures_per_line()` — consistent wrapping for TXT/PDF
- `LINES_PER_PAGE = 58`

### transpose.py — Non-Destructive Transposition

Single source of truth for the chromatic scale and enharmonic aliases.

- `CHROMATIC` — 12-note list starting at A
- `ENHARMONIC` — flat→sharp mapping (both cases: `Bb`/`bb` → `A#`)
- `transpose_symbol(symbol, semitones)` — transpose chord/note symbols
- `resolve(symbol, fret, eff) -> Resolved(symbol, fret, flag)` — resolve for display,
  with octave correction and `unplayable` flag for out-of-range frets
- `effective_transpose(doc_t, sec_t, ref_t)` — sum the three transpose levels

### songmap.py — Song Map & Riff Library Helpers

- `block_usage(doc) -> {block_id: [section_name, ...]}` — riff usage tracking
- `can_delete_block(doc, block_id) -> (bool, refs)` — guard riff deletion
- `promote_to_riff(line, sel_start, sel_end, block_id) -> (new_line, riff_items)` — Ctrl+R
- `move_section(doc, index, delta)` / `reorder_sections(doc, old, new)` — section reordering
- `chart_items(section)` / `measure_items(section)` — separate chart from tab items
- `set_chart_items(section, items)` / `set_measure_items(section, items)` — safe replacement
- `duplicate_as_reference(doc, index, new_id, new_name)` — insert section_ref copy

### export.py — TXT & PDF Export

- `build_song_lines(doc, instruments=None) -> list[str]` — TXT export
- `build_pdf(doc, instruments=None, orient="portrait") -> bytes` — PDF export
- PDF is hand-built: raw PDF operators, Helvetica + Courier fonts, zlib FlateDecode compression
- `_assemble_pdf(pages, W, H)` — page assembly ported from QLC+ Swiss Knife
- Section page-break prevention logic

### constants.py — Application Constants

- `APP_VERSION = "0.19.0"`
- `SECTION_TYPES` — list of section type labels (Verse, Chorus, Bridge, etc.)
- `INSTRUMENT_STRINGS` — dict mapping instrument names to string counts/tunings
- `RENDER_MODE_LABELS` — display labels for render modes
- `default_export_name(doc, ext)` — generate export filename from metadata

### examples.py — Built-in Example Song

- `example_document()` — returns a complete document dict with Intro, Verse, Chorus sections
  demonstrating the grammar features

## Frontend Modules

### song_writer.py — Tkinter Desktop App (~2000 lines)

The full-featured desktop frontend. Owns: chart editor bar, tab grid, song map panel, riff
library, section management, keyboard shortcuts (Ctrl+R promote, Ctrl+D duplicate-as-ref,
Alt+Up/Down reorder), transpose controls, export dialogs.

### webserver.py — Web Server (~523 lines)

- `SongStore` — `.sng` file CRUD (reads/writes from a `songs/` directory)
- `Handler` — stdlib `BaseHTTPRequestHandler` with REST API routing
- Optional `pywebview` integration via `_JsApi` class
- `ThreadingHTTPServer` with `allow_reuse_address = True`

### web/app.js — Browser Front-End (~712 lines)

- `API` object — wrapper around fetch for all REST endpoints
- State: `META`, `currentFilename`, `currentDoc`
- `fetchJSON` with error handling, `debounce`, `toast` notifications
- Section editor, chart-line input, metadata panel
