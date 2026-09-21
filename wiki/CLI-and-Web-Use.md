# CLI and Web Use

Since v0.18, the notation engine (`model.py` / `grammar.py` / `render.py`
/ `transpose.py` / `songmap.py` / `export.py`) has no Tkinter dependency,
so it can be driven from a script, a CI job, or a browser — not just the
desktop window. All three front ends read and write the same `.sng`
(plain JSON) files.

## Headless CLI — `cli.py`

No window, no Tkinter import at all.

### `convert` — one file

```bash
python3 cli.py convert -i song.sng -e pdf
python3 cli.py convert -i song.sng -e pdf -o out.pdf --orient landscape
python3 cli.py convert -i song.sng -e txt --instrument "Bass (4-string)"
```

| Flag | Meaning |
|---|---|
| `-i, --input` | path to a `.sng` file (required) |
| `-e, --export` | `txt` or `pdf` (default `pdf`) |
| `-o, --output` | output path (default: `<Artist> - <Title>.<ext>` in the current directory) |
| `--instrument` | only include this instrument's sections; repeat the flag for more than one (default: all) |
| `--orient` | `portrait` (default) or `landscape` — PDF only |

### `batch` — a whole folder

```bash
python3 cli.py batch -i songs/ -e pdf
python3 cli.py batch -i songs/ -e txt --out-dir exports/
```

Regenerates every `*.sng` in `--input-dir`. Writes alongside each `.sng`
unless `--out-dir` is given (created if it doesn't exist). Reports each
file's outcome and exits non-zero if any failed — safe to use in a CI
job or a pre-gig script.

### `lint` — sanity-check without the GUI

```bash
python3 cli.py lint -i song.sng
```

Parses every section's chart line and reports errors (or confirms they
all parse cleanly), including a re-parse round-trip check — useful for
a `.sng` file that's been hand-edited or produced by another tool. If
the song has chord shapes, it also notes chords the chart plays with no
shape and shapes for chords it never plays — as notes, never as
failures (v0.23).

Run `python3 cli.py --help` (and `... <command> --help`) for the
authoritative, always-up-to-date option list.

## Browser — `webserver.py`

A thin `http.server`-based local web server, standard library only —
same "zero external dependencies" rule as the rest of the project.

```bash
python3 webserver.py                  # opens a window automatically
python3 webserver.py --dir mysongs --port 9000
python3 webserver.py --browser        # force a normal browser tab
python3 webserver.py --no-open        # just start the server
python3 webserver.py --host 0.0.0.0   # reachable from other devices on the LAN
```

Running it opens a window automatically — no copying a URL into a
browser by hand. If the optional `pywebview` package is installed
(`pip install pywebview`, or `pip install -r requirements-optional.txt`),
that window is a native, resizable OS window — the same launch
experience as the sibling
[qlc-plus-swiss-knife-tool-script](https://github.com/giopas/qlc-plus-swiss-knife-tool-script)
project. Without it, `webserver.py` falls straight back to opening
your default browser tab; either way it's the same page underneath.
`--browser` forces the browser-tab path even with `pywebview`
installed; `--no-open` skips opening anything (useful when you're only
serving other devices on the network and don't want a window on the
host machine too).

### The native window finds its own virtualenv (v0.20)

`pywebview` is usually installed in a project venv, and the system
`python3` can't see it — so `python3 webserver.py` would silently open a
browser tab instead. The launcher now probes for a sibling virtualenv
(`.venv`, `venv`, `env`) whose interpreter can `import webview` and
re-execs into it (`os.execv` — same PID, same argv, no wrapper script,
no environment variables to set):

```bash
python3 -m venv .venv
.venv/bin/python3 -m pip install -r requirements-optional.txt
python3 webserver.py          # native window from here on
```

`SNT_NO_REEXEC=1` disables the probe; `--browser` and `--no-open` skip
it. The "am I already inside that venv?" test compares `sys.prefix` to
the venv directory rather than interpreter paths — `.venv/bin/python3`
is normally a symlink back to the base interpreter that created it, so
comparing `realpath(sys.executable)` matches in exactly the case the
re-exec exists to fix. If a venv is found but can't import `webview`,
that's reported with the underlying error rather than silently falling
back (`SNT_DEBUG_LAUNCH=1` for the full traceback).

### Where the songs are

`--dir` now defaults to your configured songs folder — initially
`~/Documents/Song Notation Tool`, *not* a folder inside the checkout.
See [[Printing and Layout]] for the config file, the one-time migration
out of the old in-repo `songs/`, and where exports land.

Open the printed URL. The front end (`web/index.html` / `app.js` /
`style.css`) gives you:

- a song list (every `.sng` in the served directory), and the songs
  folder itself with **Reveal** / **Change…** buttons — Change… can move
  your songs to the new folder (v0.25, native window only);
- a meta form (title, artist, key, time, BPM) — the song itself, and
  nothing about how it prints;
- **⚙ Layout** at the bottom of the sidebar (v0.24): a popover with every
  print setting — Size, Columns, Colour, section names, Lyrics, Recalled
  licks, Chord shapes — each with a line saying what it does, and a
  summary under the button of what's set;
- a chart-line editor per section — the same grammar as the desktop
  app's editor bar, with live parsing: type, and the fret/symbol
  preview re-renders; a parse error shows inline without clearing the
  field or the last valid render;
- a quick-insert palette beside each chart line (rests, repeat
  barlines, endings, groups, annotations, `{ tab }` and
  `{ name = tab }` licks, and one button per lick already named in the
  song — see [[Chart Line Syntax]]);
- a **Lyrics** dialog: paste, import or search; drag a section's chip
  into the whole-song sheet to drop a `=== Verse 1 ===` marker, then
  **Split by markers** (or **Split on blank lines…** for an unmarked
  sheet); **+ Section…** creates a section from there;
- a **Chords 🎸** dialog: chord shapes typed as `x32010`, a live preview
  of how they'll print, which chords are played with no shape, and
  **+ From chart** to add them;
- a per-section **Annotation** field;
- four render modes per section: Chart, Tab, Both, and **Free** (a plain
  textarea, kept and exported verbatim);
- an editable **tab grid** — add measures, type fret numbers into the
  cells, delete measures;
- add / reorder / duplicate-as-reference / delete sections;
- a collapsible "how this works" help strip, and a **Notation ⌘** panel
  documenting every mark;
- a "Start here" panel (New song / Open example) when nothing's open
  yet;
- a live Preview pane showing the full TXT export as it will look;
- TXT and PDF export, and **Print…** through the OS print dialog.

The window is an app shell: the toolbar, song list and songs-folder box
stay put and only the section list scrolls, so Save / Preview / Export
are reachable from anywhere in a long song.

**Not yet in the browser** (desktop-only for now — see the
[[Architecture]] page and `ROADMAP.md`): managing the riff library
(create/rename/delete a riff, Promote-to-riff, Duplicate-as-reference
of a *riff* rather than a section). Tab-grid editing arrived in v0.20.

### API reference

Every route lives in `webserver.py`, documented inline next to its
handler — that file is the source of truth. Summary:

| Method | Route | Does |
|---|---|---|
| `GET` | `/api/meta` | app version, section types, render modes, instrument→strings map, the labels for every ⚙ Layout choice (section layout, colour, PDF scale, columns, lyrics, recalled licks, chord shapes), and the songs + export folders |
| `GET` | `/api/songs` | list `.sng` files in the served directory |
| `GET` | `/api/songs/<name>` | load one song (migrated to the current schema), sections annotated with a transient `chart_line` |
| `PUT` | `/api/songs/<name>` | save a doc (creates the file if it doesn't exist) |
| `DELETE` | `/api/songs/<name>` | delete a `.sng` file |
| `POST` | `/api/songs` | create a new song — body `{name?, title, artist, key, time, bpm}` |
| `POST` | `/api/songs/example` | create/open the built-in example song |
| `POST` | `/api/parse` | parse one chart line — body `{line, doc?}` → items, unparsed line, rendered fret/symbol rows, or `{ok: false, error}`. With `doc` as context, references are canonicalised to ids, spelled back by name, and expanded for the rendered rows |
| `POST` | `/api/render` | render an in-memory (possibly unsaved) doc to TXT lines — body `{doc, instruments?}` |
| `POST` | `/api/lyrics/split` | cut the whole-song sheet on blank lines and propose which block goes to which section — body `{doc, text?}` |
| `POST` | `/api/lyrics/marked` | read a sheet marked with `=== Section ===` lines — body `{doc, text?}` → its segments, the section each belongs to, and names with no section yet. Writes nothing |
| `POST` | `/api/chords/shape` | parse one chord shape — body `{text, instrument?}` → frets per string, or `{ok: false, error}` |
| `POST` | `/api/chords/coverage` | chords the chart plays with no shape, shapes it never plays, and what each shape prints as under the song's transpose — body `{doc}` |
| `POST` | `/api/quit` | **⏻ Save & Close** calls this after its save: closes the native window, or stops the server in a browser tab |
| `GET` | `/api/songs/<name>/export.txt` \| `.pdf` | export a *saved* song (PDF takes `?orient=portrait\|landscape`) |
| `POST` | `/api/export.txt` \| `/api/export.pdf` | export an in-editor doc that may not be saved yet — body `{doc}` (PDF also takes `orient`) |

All bodies and responses are JSON except the two export routes, which
return the file bytes directly with a `Content-Disposition` header.

In the native window the front end also reaches Python directly through
pywebview's `window.pywebview.api` bridge, for the things a web page
can't do: `export_document` (OS save panel), `print_document` (OS print
dialog), `choose_songs_folder` (folder picker; returns what a move would
do and changes nothing), `set_songs_folder(path, move)` (switch, moving
the songs if asked), `reveal`, and `open_url`. In a browser tab that
bridge doesn't exist and the front end falls back to downloads, a new
tab, and an explanatory hint.

pywebview injects the bridge *after* the page loads, announcing it with
a `pywebviewready` event; the front end re-checks on that event, which is
why Change… and Reveal appear a moment after the window opens (v0.25).

Moving the songs folder is deliberately a bridge method and **not** an
HTTP route: the local server can't tell the app's own page from any
other page in the browser asking it to move your files, and the bridge
can only be called by the app.

### A note on exposing it beyond `localhost`

`--host 0.0.0.0` has no authentication — fine on a trusted home network
(e.g. pulling the editor up on a phone at rehearsal), but don't put it
on the open internet as-is. See `ROADMAP.md` for a planned lightweight
token.
