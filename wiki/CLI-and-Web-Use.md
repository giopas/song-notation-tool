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

### `extract`: the song out of its PDF

```bash
python3 cli.py extract -i "The Placeholders - Paper Lanterns.pdf"
```

Every PDF the app exports carries its song. `extract` writes it back out
as `Artist - Title.sng` in the current folder (or `-o` path), never over an
existing file unless you add `--force`. `convert` and `lint` also accept
such a PDF as `-i` directly.

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

- a song list (every `.sng` in the served folder) with a picker for
  **All songs**, a collection, **Hidden songs** or **Deleted songs**, and a
  **⋯** menu on each song: collections, **Hide**, **Delete…** (to the
  trash, after asking). Collections keep their own order (drag, or Alt
  with ↑ ↓) and print as one PDF. See *Collections and the trash* below;
- **Import song…** (a
  `.sng` file, or a PDF the app exported, opened as a new song in your
  songs folder), and the songs folder
  itself with **Reveal** and **Change…** buttons. **Change…** can move your
  songs to the new folder (native window only);
- a form for the song's details (title, artist, key, time, BPM), which
  says nothing about how it prints;
- **Layout** at the bottom of the sidebar: a popover with every print
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
- the **Chords** dialog: chord shapes typed as `x32010`, a live preview
  of how they will print, the chords played that have no shape, and
  **+ From chart** to add them;
- the **Riffs** dialog: create, edit, rename and delete the song's
  riffs, with where each one is used. **→ riff** beside a chart line turns
  the selected part of it into a new riff;
- a short list of suggestions as you type a chart line: riffs for a plain
  word, sections after `=`, named licks after `{`;
- **+ Part** under each section, for another instrument's line in the same
  section, as a chart line, a tab grid, both, or free text (see
  [[Printing and Layout]]);
- an Annotation field for each section;
- four render modes per section, picked with the Chart / Tab / Both / Free
  buttons in its header (Free is a plain text box, kept and exported
  exactly as typed);
- an editable tab grid: add measures, type fret numbers into the cells,
  delete measures, and make one measure longer with **+2** or **+4** (empty
  beats on every string) or shorter with **−** (empty beats only). → in
  the very last cell adds two beats. New measures can start with 3, 4, 6,
  8, 12, 16, 32 or 64 beats;
- bar lines inside a tab measure: type `|` in a cell for one after it, or
  click the gap between two cells (click it again to take it away). A
  measure's **⋯** menu puts one every 2, 3, 4, 6 or 8 beats, or takes them
  all away, and can split the measure at the selected beat, repeat it,
  copy it or duplicate it;
- selecting beats in a tab grid: Shift+click, Shift with ← or →, or drag
  across the cells; Ctrl/Cmd+A selects the whole grid. **Copy** (or
  Ctrl/Cmd+C) takes the selection, or the whole tab when nothing is
  selected, and puts it on the clipboard as text tab too; **Paste** (or
  Ctrl/Cmd+V) writes it from the cell you clicked last, over what is there,
  in this grid or any other one in the song. Ctrl/Cmd+X cuts, Delete
  empties the selected beats;
- repeats in a tab grid: **Repeat ×2** repeats the selected beats (or the
  measure the cursor is in), cutting measures where the repeat starts and
  ends. A repeat shows as **|:** before its first measure (click it to
  remove it) and **:| ×2** over its last, where the count can be changed;
- a name for a tab (**Name** beside **+ Measure**): it prints over the
  grid, and `{Name}` on any chart line plays the tab there, as a named lick
  does. The name follows the lick rules (a letter first, not starting
  like a chord). Renaming it renames every `{Name}` in the song;
- riffs named inside a tab: a measure's **⋯** menu has **Name a riff
  here**, and the name sits over that measure. `{Name}` on any chart line
  plays just that riff, so one tab (an intro of two riffs, say) can hold
  several. A riff that was a named lick before the upgrade keeps its name
  this way;
- the insert buttons under the chart line you are typing in, in four
  groups: Lines, Repeats, Licks and Marks;
- a mistake in a chart line explained under it in plain words, with the
  chart that last worked kept on screen, dimmed;
- adding, moving, copying and deleting sections. Drag a section by its
  **⠿** handle to move it: the cards fold down to their names while you
  drag, so the whole song is on screen, and Escape cancels. With the
  handle clicked, the up and down arrow keys move it one place. The
  card's **⋯** menu holds its print size, Delete, **Copy as a
  reference…** and **Duplicate as a copy…**. Both open a short list of
  every place the copy can go (the top, or after any section). A reference
  plays the original, with every part (each one shows which part it
  plays, with **Make it its own copy**) and, for a tab section, its tab
  grid. A duplicate is independent: its chart line, tab, parts and text
  are its own;
- a "how this works" help strip you can fold away, and the **Notation**
  panel, which explains every mark;
- a "Start here" panel (New song, Open example) when nothing is open;
- a live Preview beside the song, so Save and Export stay in reach: **Text**
  shows the whole TXT export as you type, **PDF** shows the PDF as it will
  print (portrait or landscape) a moment after you stop typing, and the
  open button beside it opens that PDF in your browser or PDF viewer;
- an older song whose tab was typed as text opens with that tab in the
  grid: a Free section (or part) of text tab (`G|--5--7--|` rows), or a
  chart line of nothing but licks and rests (a rest becomes an empty
  measure, a bar long; each named lick keeps its name on its measures),
  becomes a tab grid with its bars and
  repeats (`|:`, `:|`, an `x4` line) and its name. It happens once per
  song, when it is opened; the file changes when you save it, and the
  Free text stays in the section, so switching back to Free shows it;
- dialogs that close with Escape and keep the keyboard inside them while
  open; song tools that stay greyed out until a song is open;
- a light and a dark theme (**◐**; the system's until you choose), and a
  narrow layout for a phone, with the song list behind **☰**;
- TXT and PDF export, and **Print…** through the system print dialog.

The toolbar, the song list and the songs folder box stay in place and only
the section list scrolls, so **Save**, **Preview** and **Export** are
always in reach in a long song.

Not yet in the browser (desktop only for now; see [[Architecture]] and
`ROADMAP.md`): duplicating a riff (as opposed to a section) as a
reference. The other way round, parts, the riff dialog and the
suggestions while typing are in the browser app only.

### API reference

Every route is in `webserver.py`, documented next to its handler. That
file is the source of truth. In short:

| Method | Route | Does |
|---|---|---|
| `GET` | `/api/meta` | app version, section types, render modes, the instrument-to-strings map, the labels for every Layout choice (section layout, colour, PDF scale, columns, lyrics, recalled licks, chord shapes), and the songs and export folders |
| `GET` | `/api/songs` | list the `.sng` files in the served folder |
| `GET` | `/api/songs/<name>` | load one song (migrated to the current schema), with a temporary `chart_line` added to each section |
| `PUT` | `/api/songs/<name>` | save a song (creating the file if needed) under its own name, `Artist - Title.sng`, renaming the file if the title or artist changed → `{ok, filename, renamed_from}`; use `filename` from then on |
| `DELETE` | `/api/songs/<name>` | move a song to the trash (`Deleted songs` in the songs folder) → `{ok, trashed_as}` |
| `GET` | `/api/library` | collections (each `{id, name, songs}` in order), hidden songs and `trash_count` |
| `POST` | `/api/library` | body `{op, ...}`: `hide` (`song`, `hidden`), `create` (`name`, `songs?`), `rename` (`id`, `name`), `delete` (`id`; the songs stay), `add` (`id`, `songs`), `remove` (`id`, `song`), `move` (`id`, `song`, `to`), `order` (`id`, `songs`) → `{ok, library}`, or `{ok: false, error}`. JSON only |
| `GET` | `/api/trash` | the deleted songs, newest first: `{name, original, title, artist, deleted_at}` |
| `POST` | `/api/trash` | body `{op, name?}`: `restore` → `{filename}`, `purge` (one song for good), `empty` → `{deleted}`. JSON only |
| `GET` | `/api/collections/<id>/export.pdf` | every song of a collection in one PDF, in order (`?orient=portrait\|landscape`); `422` for an empty collection or a song that fails its check |
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
| `POST` | `/api/riffs` | the riff library, body `{op, doc, ...}` with `op` one of `create` (`name`, `line`), `set_line` (`id`, `line`), `rename` (`id`, `name`), `delete` (`id`), `promote` (`section_id`, `line`, `start`, `end`, `name?`) or `list` → `{ok, doc, riffs: [{id, line, used_in}]}`, or `{ok: false, error}`. Changes only the song it is sent; the editor saves it as usual |
| `POST` | `/api/import` | import a song, body `{filename, data}` with the file in base64: a `.sng`, or a PDF the app exported. Always saved as a new file named after the song → `{ok, filename, doc}`, or `{ok: false, error}` |
| `POST` | `/api/quit` | called by **⏻** (Save and close) after it saves: closes the native window, or stops the server in a browser tab |
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
`export_document` (system save panel), `export_collection(id, orient)`
(the same for a collection's PDF), `print_document` (system print
dialog), `choose_songs_folder` (folder picker; returns what a move would
do and changes nothing), `set_songs_folder(path, move)` (switch, moving
the songs if asked), `reveal`, and `open_url`. In a browser tab there is
no bridge, and the front end uses downloads, a new tab and a hint
instead.

pywebview adds the bridge after the page has loaded and announces it with
a `pywebviewready` event. The front end checks again on that event, which
is why **Change…** and **Reveal** appear a moment after the window opens.

`/api/library` and `/api/trash` accept only `Content-Type:
application/json`. Another site's page can send a form or plain-text POST
to the local server without the browser asking first, but not a JSON one,
so it cannot empty your trash.

### Collections and the trash

Hidden songs and collections live in `library.json` in the songs folder,
and deleted songs in its `Deleted songs` folder until the trash is
emptied. No song file changes when it is hidden or added to a collection.
A name in `library.json` whose file has gone (deleted or renamed outside
the app) is skipped, not an error. Saving a song under a new name updates
its collections, and **Change…** takes `library.json` and the deleted
songs along when it moves the songs; a `library.json` already in the new
folder is kept, and ours stays in the old one.

Moving the songs folder is a bridge method on purpose, and not an HTTP
route. The local server cannot tell the app's own page from any other
page in the browser asking it to move your files, while the bridge can
only be called by the app.

### Using it beyond `localhost`

`--host 0.0.0.0` has no authentication. That is fine on a trusted home
network (opening the editor on a phone at rehearsal, say), but do not put
it on the open internet as it is. `ROADMAP.md` has a planned lightweight
token.
