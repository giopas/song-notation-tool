# Architecture

## The engine

These modules have no Tkinter import, and that stays so. It is what lets
three front ends share them. `tests/test_export.py` checks it for
`export.py`.

| Module | What it owns |
|---|---|
| `model.py` | The schema, constructors for items, sections and blocks, reading and writing `.sng`, and the migration from v0.15 files |
| `grammar.py` | The chart-line grammar: `parse()`, `parse_items()`, `unparse()`. See [[Chart Line Syntax]] |
| `transpose.py` | Transposing notes and symbols, and working out the offset at render time |
| `render.py` | Row rendering, mark spans and the page-count estimate, used by the song map, Stage View, TXT and PDF. `resolve_references()` expands a section or riff reference, or a `{Riff1}` lick reference, into what it plays. `chart_body_lines()` lays out a section's body and passes free text through as typed. `render_chart_rows()` and `chart_body_rows()` return each row as `{text, role, spans}`, so the PDF and the browser colour licks and rests from one source. Lyric placement: `lyrics_mode()`, `section_lyric_lines()`, `lyric_blocks()` and `gathered_lyrics()` for the "all together" layouts, and `compose_beside()` and `shared_lyric_column()` for words beside the chart |
| `lyrics.py` | The lyric sheet: splitting on blank lines and proposing which block goes to which section, and the `=== Section ===` markers (`split_marked()`, `match_sections()`, `apply_marked()`, `strip_markers()`) |
| `chords.py` | Chord shapes: parsing `x32010`, drawing a shape as a small tab diagram, laying the shapes out in rows, checking them against the chords the chart plays (`coverage()`), and re-voicing a shape for a transpose (`transpose_chord()`: open to open, otherwise an E-form or A-form barre; movable shapes slide) |
| `songmap.py` | Song-map operations: reorder, promote to riff, duplicate as reference, riff usage, and splitting chart items from measure items. Reference lookup and the split between what is stored and what is shown (`find_section()`, `display_ref_key()`, `display_items()`, `canonicalise_refs()`): ids in the file, names on screen. The named-lick index, worked out from the song each time instead of stored (`lick_index()`, `find_lick()`, `lick_names()`) |
| `export.py` | The TXT and PDF builders, `build_song_lines(doc, instruments=None)` and `build_pdf(doc, instruments=None, orient=..., date=None)`, and the one writer every front end uses: `export_to_file()` checks the song (`check_document()`, which `cli.py lint` also runs), builds the bytes (`export_bytes()`) and writes them (`write_export()`: atomic, and never over an existing file unless the caller says the user agreed). `prepare_export()` does the check and the bytes for the browser's downloads. `build_pdf` resolves the song's size setting and calls `_build_pdf(..., scale)`, which also returns the page count, so `resolve_scale()` can search for the largest scale that still fits. `date` sets the footer date (today by default), so the same song and date give the same bytes. See [[Printing and Layout]] |
| `examples.py` | `example_document()`, the song behind **Open example song** |
| `userpaths.py` | Where the user's files live: the songs folder, the last export folder, the config file for each platform, and the one-time move out of the old in-repo `songs/`. Changing the songs folder: `plan_relocation()` says what a move would do without doing it, and `relocate_songs()` moves the `.sng` files, never overwriting, then switches |
| `update.py` | The update check and **Update and restart** for the packaged app (adapted from QLC+ Swiss Knife): asking the GitHub Releases API, picking this computer's download, checking it against `SHA256SUMS`, unpacking it safely and writing the script that swaps it in after the app closes. The only network code in the app. See [[Installing and Updating]] |
| `constants.py` | `APP_VERSION`, the instrument-to-strings map, section types, render-mode labels, the labels for every **⚙ Layout** choice, and `default_export_name()` |

Each of them takes and returns plain dicts and lists shaped like
`model.py`'s schema. None of them holds a widget or any live UI state.

## The three front ends

| Front end | File | What it adds to the engine |
|---|---|---|
| Desktop | `song_writer.py` | `SongNotationApp` (Tkinter): the song map, the chart editor bar, the riff strip, the tab grid, Stage View, the Start Here panel, the help strip, the live Preview window, the Lyrics dialog's marker buttons, the Chords dialog, and the **⚙ Layout** bar along the bottom of the window |
| Command line | `cli.py` | `convert`, `batch` and `lint` subcommands built on `argparse`, with no window |
| Browser | `webserver.py` and `web/` | A JSON API on the standard library's `http.server`, and a plain HTML, CSS and JavaScript front end with no build step and no framework. In the native window, pywebview's `js_api` bridge (`_JsApi` in `webserver.py`) adds what a web page cannot do: save panels, printing, Finder, and changing or moving the songs folder |

None of the three has its own copy of the TXT and PDF builders, the chart
grammar or the transposition code. They all call the same engine modules.

## The packages

`packaging/songnotation.spec` builds the native window with PyInstaller:
`webserver.py` as the program, `web/` as data, and the icons from
`packaging/icons/`. In the package `webserver.WEB_DIR` points into the
bundle (`sys._MEIPASS`), the virtualenv switch and the old `songs/` move
are skipped, and the previous version is cleaned up after an update.
`packaging/make_archive.py` turns the build into the release files, with
the names `update.py` looks for; `packaging/windows/installer.iss` is the
Inno Setup installer, and `packaging/linux/` holds the menu entry. The
release workflow (`.github/workflows/release.yml`) runs all of this on a
version tag; see [[Installing and Updating]].

## Where a change goes

- A rendering or export bug or feature (a misaligned column, a new export
  field, a new mark type): `render.py` and/or `export.py`. All three front
  ends get it.
- A change to the chart-line syntax: `grammar.py`, then update
  [[Chart Line Syntax]].
- A song-map operation (a new way to reorder, duplicate or reference):
  `songmap.py`, then connect it in the front ends that should offer it. Not
  every front end has to; see the list of what is not yet in the browser
  on [[CLI and Web Use]].
- A desktop-only change (a toolbar button, a dialog): `song_writer.py`
  only.
- A command-line-only change (a subcommand, an output format): `cli.py`
  only.
- A web-only change (an API route, a front-end feature): `webserver.py`
  and `web/` only.
- A constant used by more than one front end (the instrument list, section
  types, the app version): `constants.py`.

## Testing

`tests/` covers the engine modules with plain pytest, and needs no Tkinter
to run. The command line is tested through a real subprocess, and the web
server through its real HTTP routes.

`tests/golden/` holds ten invented songs, one per layout feature, with the
TXT and PDF each one must produce (`tests/test_golden.py`). TXT is compared
byte for byte; PDF is compared with its page streams inflated, because
different zlib builds pack the same drawing into different bytes. The
footer date and the version are fixed for these files, so a release does
not change them. When an output change is intended, run
`python3 tools/make_golden.py`, look at the new PDFs, and commit them with
the change. The desktop app is smoke-tested with Tkinter running under Xvfb:
create a `SongNotationApp`, drive it through `_start_open_example()`,
`_toggle_help()`, `_toggle_preview()`, `_build_song_lines()` and
`_build_pdf()`, and check the resulting state. There is no
`tests/test_song_writer.py` yet. It would be worth adding if the desktop
UI keeps growing.
