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
| `render.py` | shared row rendering, mark-span resolution, page-count estimate — used by the song map, Stage View, TXT, and PDF alike. Since v0.20 also `resolve_references()` (expand a section/riff reference into what it actually plays) and `chart_body_lines()` (lay out a section's body, passing free text through un-aligned) |
| `songmap.py` | song-map operations: reorder, promote-to-riff, duplicate-as-reference, riff usage tracking, chart-vs-measure item splitting. Since v0.20 also reference lookup and the display/storage split — `find_section()`, `display_ref_key()`, `display_items()`, `canonicalise_refs()`: ids in the file, names on screen |
| `export.py` | `build_song_lines(doc, instruments=None)` and `build_pdf(doc, instruments=None, orient=...)` — the actual TXT/PDF builders, added in v0.18. Since v0.20 `build_pdf` resolves the document's size setting and delegates to `_build_pdf(..., scale)`, which returns the page count too so `resolve_scale()` can search for the largest scale that still fits. See [[Printing and Layout]] |
| `examples.py` | `example_document()` — the "Open example" sample, added in v0.18 |
| `userpaths.py` | where the user's own files live — the songs folder, the last-used export folder, the per-platform config file, and the one-time migration out of the old in-repo `songs/`. Added in v0.20 |
| `constants.py` | `APP_VERSION`, instrument→strings map, section types, render-mode labels, `default_export_name()`, added in v0.18 |

Every one of these takes and returns plain dicts/lists shaped like
`model.py`'s schema — no widget, no live UI state, anywhere.

## The three front ends

| Front end | File | What it adds over the engine |
|---|---|---|
| Desktop | `song_writer.py` | `SongNotationApp` (Tkinter): the song map, chart editor bar, riff strip, tab grid, Stage View, and — since v0.18 — the Start Here panel, help strip, and live Preview window |
| Headless | `cli.py` | `argparse`-based `convert` / `batch` / `lint` subcommands, no window |
| Browser | `webserver.py` + `web/` | stdlib `http.server` JSON API + a vanilla HTML/CSS/JS front end (no build step, no framework) |

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
