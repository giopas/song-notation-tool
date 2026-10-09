# CLI and web use

The notation engine (`model.py`, `grammar.py`, `render.py`,
`transpose.py`, `songmap.py`, `export.py`) does not depend on Tkinter, so
a script, a CI job or a browser can drive it as well as the desktop
window. All three front ends read and write the same `.sng` files (plain
JSON).

## Command line: `cli.py`

No window, and no Tkinter import at all.

### `convert`: one file

```bash
python3 cli.py convert -i song.sng -e pdf
python3 cli.py convert -i song.sng -e pdf -o out.pdf --orient landscape
python3 cli.py convert -i song.sng -e txt --instrument "Bass (4-string)"
```

| Flag | Meaning |
|---|---|
| `-i, --input` | path to a `.sng` file (required) |
| `-e, --export` | `txt` or `pdf` (default `pdf`) |
| `-o, --output` | output path (default: `<Artist> - <Title>.<ext>` in the current folder) |
| `--instrument` | only this instrument's sections; repeat the flag for more than one (default: all) |
| `--orient` | `portrait` (default) or `landscape`, PDF only |
| `--force` | replace the output file if it exists (default: keep it and write to the next free name, `song_v2.pdf`) |

An existing file is never replaced without `--force`. The export goes to
the next free name instead (`out.pdf` becomes `out_v2.pdf`, then
`out_v3.pdf`; a name that already ends in `_vN` counts on from N), and
`convert` says which file it kept. Before writing anything, `convert` runs
the same check as `lint`: a song that fails it is not exported, the errors
go to standard error, and the exit code is 1. Its notes are printed and
never stop the export.

### `batch`: a whole folder

```bash
python3 cli.py batch -i songs/ -e pdf
python3 cli.py batch -i songs/ -e txt --out-dir exports/
```

Regenerates every `*.sng` in `--input-dir`. The files are written next to
each `.sng` unless you give `--out-dir`, which is created if it does not
exist. It takes the same `--force` as `convert`: without it, running the
batch again writes `song_v2.pdf` next to last time's `song.pdf`. It
checks each song first and reports what happened to each file (written,
not exported because of the check, or failed), and exits with a non-zero
code if any was not written, so you can use it in a CI job or a script
before a gig.

### `lint`: check a song without the GUI

```bash
python3 cli.py lint -i song.sng
```

Parses every section's chart line and reports the errors, or confirms
that all of them parse, including a check that each line parses again
the same after being written back out. That helps with a `.sng` file
edited by hand or made by another tool. If the song has chord shapes, it
also lists chords the chart plays that have no shape, and shapes for
chords it never plays. These are notes, never failures.

This is the check every export runs first, in every front end
(`export.check_document()`), so `lint` tells you in advance whether an
export will go through.

`python3 cli.py --help` (and `... <command> --help`) always has the
current option list.

## Browser: `webserver.py`

A small local web server built on `http.server`, standard library only,
like the rest of the project.

```bash
python3 webserver.py                  # opens a window automatically
python3 webserver.py --dir mysongs --port 9000
python3 webserver.py --browser        # use a normal browser tab
python3 webserver.py --no-open        # just start the server
python3 webserver.py --host 0.0.0.0   # reachable from other devices on the LAN
python3 webserver.py --smoke          # check the page, files and an export, then exit
```

Without `--port` it listens on 8420, or on any free port if 8420 is taken.

It opens a window by itself, so there is no URL to copy into a browser.
If the optional `pywebview` package is installed (`pip install pywebview`,
or `pip install -r requirements-optional.txt`), the window is a native,
resizable window, the same as in the sibling
[qlc-plus-swiss-knife-tool-script](https://github.com/giopas/qlc-plus-swiss-knife-tool-script)
project. Without it, `webserver.py` opens your default browser instead.
The page is the same either way. `--browser` uses a browser tab even when
`pywebview` is installed. `--no-open` opens nothing, which is handy when
you only serve other devices on the network and do not want a window on
the computer running the server.

### The native window finds its own virtualenv

`pywebview` is usually installed in a project virtualenv that the system
`python3` cannot see, so `python3 webserver.py` would quietly open a
browser tab. To avoid that, the launcher looks for a virtualenv next to it
(`.venv`, `venv`, `env`) whose interpreter can `import webview`, and
restarts itself in it with `os.execv` (same process id, same arguments, no
wrapper script, no environment variables to set):

```bash
python3 -m venv .venv
.venv/bin/python3 -m pip install -r requirements-optional.txt
python3 webserver.py          # native window from now on
```

`SNT_NO_REEXEC=1` turns this off, and `--browser` and `--no-open` skip it.
To tell whether it is already running inside that virtualenv, it compares
`sys.prefix` with the virtualenv folder. Comparing interpreter paths would
not work: `.venv/bin/python3` is normally a symlink to the interpreter that
created it, so `realpath(sys.executable)` would match in exactly the case
the restart is meant to fix. If it finds a virtualenv that cannot import
`webview`, it reports the error instead of quietly falling back
(`SNT_DEBUG_LAUNCH=1` prints the full traceback).

### Where the songs are

`--dir` defaults to your configured songs folder, which starts as
`~/Documents/Song Notation Tool`, outside the checkout. [[Printing and Layout]]
covers the config file, the one-time move out of the old in-repo
`songs/` folder, and where exports go.

Open the URL it prints. The front end (`web/index.html`, `app.js`,
`style.css`) has:

- a song list (every `.sng` in the served folder), and the songs folder
  itself with **Reveal** and **Change…** buttons. **Change…** can move your
  songs to the new folder (native window only);
- a form for the song's details (title, artist, key, time, BPM), which
  says nothing about how it prints;
- **⚙ Layout** at the bottom of the sidebar: a popover with every print
  setting (Size, Columns, Colour, section names, Lyrics, Recalled licks,
  Chord shapes), each with a line saying what it does, and a summary of
  what is set under the button;
- a chart-line editor for each section, with the same grammar as the
  desktop app's editor bar. As you type, the fret and symbol preview
  updates. A parse error shows in place without clearing the field or the
  last good preview;
- a quick-insert palette beside each chart line: rests, repeat barlines,
  endings, groups, annotations, `{ tab }` and `{ name = tab }` licks, and
  one button for each lick already named in the song (see
  [[Chart Line Syntax]]);
- the **Lyrics** dialog: paste, import or search; drag a section's chip
  into the whole-song sheet to drop a `=== Verse 1 ===` marker, then
  **Split by markers** (or **Split on blank lines…** for a sheet without
  markers). **+ Section…** creates a section from there;
- the **Chords 🎸** dialog: chord shapes typed as `x32010`, a live preview
  of how they will print, the chords played that have no shape, and
  **+ From chart** to add them;
- an Annotation field for each section;
- four render modes per section: Chart, Tab, Both and Free (a plain text
  box, kept and exported exactly as typed);
- an editable tab grid: add measures, type fret numbers into the cells,
  delete measures;
- adding, reordering, duplicating as reference and deleting sections;
- a "how this works" help strip you can fold away, and the **Notation ⌘**
  panel, which explains every mark;
- a "Start here" panel (New song, Open example) when nothing is open;
- a live Preview showing the whole TXT export as it will look;
- TXT and PDF export, and **Print…** through the system print dialog.

The toolbar, the song list and the songs folder box stay in place and only
the section list scrolls, so **Save**, **Preview** and **Export** are
always in reach in a long song.

Not yet in the browser (desktop only for now; see [[Architecture]] and
`ROADMAP.md`): managing the riff library, which means creating, renaming
or deleting a riff, **Promote to riff**, and duplicating a riff (as opposed
to a section) as a reference.

### API reference

Every route is in `webserver.py`, documented next to its handler. That
file is the source of truth. In short:

| Method | Route | Does |
|---|---|---|
| `GET` | `/api/meta` | app version, section types, render modes, the instrument-to-strings map, the labels for every ⚙ Layout choice (section layout, colour, PDF scale, columns, lyrics, recalled licks, chord shapes), and the songs and export folders |
| `GET` | `/api/songs` | list the `.sng` files in the served folder |
| `GET` | `/api/songs/<name>` | load one song (migrated to the current schema), with a temporary `chart_line` added to each section |
| `PUT` | `/api/songs/<name>` | save a song (creating the file if needed) under its own name, `Artist - Title.sng`, renaming the file if the title or artist changed → `{ok, filename, renamed_from}`; use `filename` from then on |
| `DELETE` | `/api/songs/<name>` | delete a `.sng` file |
| `POST` | `/api/songs` | create a song, body `{name?, title, artist, key, time, bpm}`; named `Artist - Title.sng` unless `name` is given, with "(2)" if the name is taken |
| `POST` | `/api/songs/example` | open the built-in example as shipped: an untouched copy is reused, otherwise a fresh one is created, never an example you have since edited into a song |
| `POST` | `/api/parse` | parse one chart line, body `{line, doc?}` → items, the line written back out, rendered fret and symbol rows, or `{ok: false, error}`. With `doc` as context, references are stored as ids, shown by name, and expanded in the rendered rows |
| `POST` | `/api/render` | render a song held in memory (possibly unsaved) to TXT lines, body `{doc, instruments?}` |
| `POST` | `/api/lyrics/split` | cut the whole-song sheet at blank lines and propose which block goes to which section, body `{doc, text?}` |
| `POST` | `/api/lyrics/marked` | read a sheet marked with `=== Section ===` lines, body `{doc, text?}` → its segments, the section each belongs to, and names with no section yet. Writes nothing |
| `POST` | `/api/chords/shape` | parse one chord shape, body `{text, instrument?}` → frets per string, or `{ok: false, error}` |
| `POST` | `/api/chords/coverage` | chords the chart plays with no shape, shapes it never plays, and what each shape prints as under the song's transpose, body `{doc}` |
| `GET` | `/api/update/check` | the update check (`?force=1` asks again instead of using the answer from the last 12 hours): current and latest version, whether it is newer, the release notes, and whether this copy can install it. See [[Installing and Updating]] |
| `POST` | `/api/update/settings` | body `{check}`: switch the check at start on or off |
| `POST` | `/api/update/install` | installed app only: download, verify and stage the new version, then close so it can be swapped in; from the sources a `400` that says to use `git pull` |
| `POST` | `/api/update/open` | open the release page in the system browser (only the project's own pages) |
| `POST` | `/api/quit` | called by **⏻ Save & Close** after it saves: closes the native window, or stops the server in a browser tab |
| `GET` | `/api/songs/<name>/export.txt` \| `.pdf` | export a saved song (PDF takes `?orient=portrait\|landscape`) |
| `POST` | `/api/export.txt` \| `/api/export.pdf` | export a song from the editor that may not be saved yet, body `{doc}` (PDF also takes `orient`) |

Bodies and responses are JSON, except the two export routes, which return
the file itself with a `Content-Disposition` header. The export routes
check the song first. A song that fails the check gets a `422` with
`{ok: false, error, errors}` and no file; the check's notes, if any, come
with the file in an `X-Export-Warnings` header (a JSON list), and the
front end shows them after the download.

In the native window the front end also calls Python directly through
pywebview's `window.pywebview.api` bridge, for what a web page cannot do:
`export_document` (system save panel), `print_document` (system print
dialog), `choose_songs_folder` (folder picker; returns what a move would
do and changes nothing), `set_songs_folder(path, move)` (switch, moving
the songs if asked), `reveal`, and `open_url`. In a browser tab there is
no bridge, and the front end uses downloads, a new tab and a hint
instead.

pywebview adds the bridge after the page has loaded and announces it with
a `pywebviewready` event. The front end checks again on that event, which
is why **Change…** and **Reveal** appear a moment after the window opens.

Moving the songs folder is a bridge method on purpose, and not an HTTP
route. The local server cannot tell the app's own page from any other
page in the browser asking it to move your files, while the bridge can
only be called by the app.

### Using it beyond `localhost`

`--host 0.0.0.0` has no authentication. That is fine on a trusted home
network (opening the editor on a phone at rehearsal, say), but do not put
it on the open internet as it is. `ROADMAP.md` has a planned lightweight
token.
