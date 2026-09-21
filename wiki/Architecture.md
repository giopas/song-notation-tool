# Architecture

## The pure engine

These modules have **no Tkinter import**, and never will — that's the
whole point of the v0.18 split (`tests/test_export.py` asserts it
explicitly for `export.py`; the others were already this way as of
v0.16/v0.17):

| Module | Owns |
|---|---|
| `model.py` | schema, item/section/block constructors, `.sng` (de)serialisation, v0.15→v0.16 migration |
| `grammar.py` | the chart-line grammar — `parse()` / `parse_items()` / `unparse()`. See [[Chart Line Syntax]] |
| `transpose.py` | note/symbol transposition math, render-time offset resolution |
| `render.py` | shared row rendering, mark-span resolution, page-count estimate — used by the song map, Stage View, TXT, and PDF alike. Since v0.20 also `resolve_references()` (expand a section/riff reference into what it actually plays — and, since v0.23, a `{Riff1}` lick reference into its lick) and `chart_body_lines()` (lay out a section's body, passing free text through un-aligned). Since v0.21 also `render_chart_rows()` / `chart_body_rows()`, which return each row as `{text, role, spans}` so the PDF and the browser colour licks and rests from one source rather than re-deriving them from the text. Since v0.24, lyric placement: `lyrics_mode()`, `section_lyric_lines()`, `lyric_blocks()` / `gathered_lyrics()` for the all-together layouts, `compose_beside()` and `shared_lyric_column()` for words beside the chart |
| `lyrics.py` | the lyric sheet: splitting on blank lines and proposing which block goes to which section (v0.22), and since v0.23 `=== Section ===` markers — `split_marked()`, `match_sections()`, `apply_marked()`, `strip_markers()` |
| `chords.py` | chord shapes (v0.23): parsing `x32010`, drawing a shape as a small tab diagram, laying the chord sheet out in rows, checking shapes against the chords the chart plays (`coverage()`), and re-voicing a shape for a transpose (`transpose_chord()` — open to open, else E/A-form barre, movable shapes slide) |
| `songmap.py` | song-map operations: reorder, promote-to-riff, duplicate-as-reference, riff usage tracking, chart-vs-measure item splitting. Since v0.20 also reference lookup and the display/storage split — `find_section()`, `display_ref_key()`, `display_items()`, `canonicalise_refs()`: ids in the file, names on screen. Since v0.23 the named-lick index, derived by walking the song rather than stored — `lick_index()`, `find_lick()`, `lick_names()` |
| `export.py` | `build_song_lines(doc, instruments=None)` and `build_pdf(doc, instruments=None, orient=...)` — the actual TXT/PDF builders, added in v0.18. Since v0.20 `build_pdf` resolves the document's size setting and delegates to `_build_pdf(..., scale)`, which returns the page count too so `resolve_scale()` can search for the largest scale that still fits. See [[Printing and Layout]] |
| `examples.py` | `example_document()` — the "Open example" sample, added in v0.18 |
| `userpaths.py` | where the user's own files live — the songs folder, the last-used export folder, the per-platform config file, and the one-time migration out of the old in-repo `songs/`. Added in v0.20. Since v0.25 also changing the songs folder: `plan_relocation()` (what a move would do, without doing it) and `relocate_songs()` (move the `.sng` files, never overwriting, then switch) |
| `constants.py` | `APP_VERSION`, instrument→strings map, section types, render-mode labels, the labels for every ⚙ Layout choice, `default_export_name()`, added in v0.18 |

Every one of these takes and returns plain dicts/lists shaped like
`model.py`'s schema — no widget, no live UI state, anywhere.

## The three front ends

| Front end | File | What it adds over the engine |
|---|---|---|
| Desktop | `song_writer.py` | `SongNotationApp` (Tkinter): the song map, chart editor bar, riff strip, tab grid, Stage View, and — since v0.18 — the Start Here panel, help strip, and live Preview window. Since v0.23 the Lyrics dialog's marker buttons and the Chords dialog; since v0.24 the ⚙ Layout bar along the bottom of the window |
| Headless | `cli.py` | `argparse`-based `convert` / `batch` / `lint` subcommands, no window |
| Browser | `webserver.py` + `web/` | stdlib `http.server` JSON API + a vanilla HTML/CSS/JS front end (no build step, no framework). In the native window, pywebview's `js_api` bridge (`_JsApi` in `webserver.py`) adds what a web page can't do — save panels, printing, Finder, and changing/moving the songs folder |

None of the three duplicates the TXT/PDF builder, the chart grammar, or
the transposition math — they all call into the same engine module.
Before v0.18, `song_writer.py` had its own copies of the TXT/PDF
builders reading straight from live Tkinter `StringVar`s; that's what
`export.py` replaced (see the v0.18.0 entry in `CHANGELOG.md`).

## Where a change goes

- **A rendering or export bug/feature** (a column misaligned, a new
  export field, a new mark type) → `render.py` and/or `export.py`. All
  three front ends pick it up automatically.
- **A chart-line syntax change** → `grammar.py`, then update
  [[Chart Line Syntax]].
- **A song-map operation** (a new way to reorder/duplicate/reference)
  → `songmap.py`, then wire it into whichever front end(s) should
  expose it — not all three necessarily need to, see the "not yet in
  the browser" note on [[CLI and Web Use]].
- **A desktop-only UI change** (a new toolbar button, a dialog) →
  `song_writer.py` only.
- **A CLI-only concern** (a new subcommand, output format) → `cli.py`
  only.
- **A web-only concern** (a new API route, a front-end feature) →
  `webserver.py` and `web/` only.
- **A constant shared by more than one front end** (instrument list,
  section types, the app version) → `constants.py`, imported by
  whichever front ends need it.

## Testing

`tests/` covers the engine modules with plain pytest (no Tkinter
needed to run them). The desktop app itself is smoke-tested with
Tkinter running under Xvfb in CI-style checks (instantiate
`SongNotationApp`, drive it through `_start_open_example()` /
`_toggle_help()` / `_toggle_preview()` / `_build_song_lines()` /
`_build_pdf()`, assert on the resulting state) — there's no dedicated
`tests/test_song_writer.py` yet; that's a reasonable next addition if
the desktop UI grows much further.
