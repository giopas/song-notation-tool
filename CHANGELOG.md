# Changelog

All notable changes to this project are documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.0.0/).

## [0.25.0] — 2026-09-21

### Added
- **Change the songs folder — and take your songs with you.** Change… in
  the sidebar opens the system folder picker, then says what would happen
  before anything does: both paths, how many songs will move, and which
  ones will stay because a song with the same name is already there.
  Three choices: **Move N songs**, **Just use this folder** (point the app
  there, leave the songs where they are), or **Cancel**.

  The move is deliberately conservative: only `.sng` files move (the
  folder could be ~/Documents itself, and nothing else in it is ours), a
  song already in the new folder is never overwritten, the old folder is
  never deleted, and it works across drives and into iCloud Drive. The
  setting only changes once the new folder is known to be usable. The song
  you have open stays open, unsaved edits and all, if it moved with the
  rest; if it didn't, it's closed rather than left for a later Save to
  write a stray copy into the new folder.

  Native window only, on purpose: the local web server has no way to tell
  the app's own page from any other page in the browser asking it to move
  files, so the move is reachable only through pywebview's bridge. In a
  browser tab, `--dir` still chooses the folder per run.

### Fixed
- **Change… and Reveal were missing in the app's own window**, which
  showed the browser-tab hint instead. The page checked for pywebview's
  bridge at startup, before pywebview had injected it; it now also checks
  when the `pywebviewready` event says the bridge has arrived.

### Documentation
- README and wiki brought in line with everything from v0.23 on: the
  ⚙ Layout menu and all seven settings, the six lyric placements and the
  shared lyric column, section markers, named licks (and that braces —
  not `=` — recall one), recalled-lick printing, chord shapes with the
  coverage check and transpose re-voicing, tab drawn on the column grid,
  and changing/moving the songs folder. The wiki's API table gains the
  lyrics, chords and quit routes; Architecture lists `lyrics.py` and
  `chords.py`.
- Corrected: the wiki said the PDF body is set in Courier. It's
  Helvetica positioned on a Courier-width column grid; and the README
  said references aren't expanded in the export, which stopped being
  true in v0.20.

### API
- `userpaths.py`: `song_files()`, `plan_relocation()`, `relocate_songs()`.
- Native bridge: `choose_songs_folder()` now returns a plan and changes
  nothing; `set_songs_folder(path, move)` does the switch.

## [0.24.3] — 2026-09-21

### Fixed
- **Lick tab lines line up in the PDF.** Each string line was drawn as one
  piece of text in a proportional face, where a "4" is wider than a "-",
  so a line with more digits came out longer than its neighbours — and a
  fret on one string drifted off the fret it's played with on the next.
  Every mark is now placed on its own column of the grid the chords use,
  and each run of dashes is drawn as a solid rule from mark to mark, so a
  string reads as one line between its bar lines. TXT and the on-screen
  preview were already right; only the PDF changes.

## [0.24.2] — 2026-09-21

### Changed
- **A named lick's name sits in front of its tab**, on the tab's first
  string line — `Riff1 |D|-----3---|` — instead of on a line of its own
  above it, where it read as a separate line of the chart. The other
  string lines are indented to match so the tab stays a grid, and a
  repeat count comes along: `Riff1 (x2) |D|…`. Applies to the lick where
  it's defined and wherever `{Riff1}` recalls it as tab; "name only"
  recalls still print on the chord row, since there's no tab to sit by.

## [0.24.1] — 2026-09-21

### Fixed
- **Lyrics beside the chart now line up down the page.** Each section used
  to start its words just past its *own* chart, so a narrow chorus put them
  in one column and a wide verse in another, and down the page they
  zigzagged. Every section's words now start in one column for the whole
  song, clear of the widest chart that has words beside it; a wide
  instrumental section with no words doesn't push them out.
  `render.shared_lyric_column()`, used by the TXT, the PDF and the width fit.

## [0.24.0] — 2026-09-21

Where the words go is now one choice for the whole song, made on the main
screen — and every print setting has moved into one Layout menu at the
bottom left.

### Added
- **Six lyric placements, one setting.** ⚙ Layout → Lyrics:
  - **Don't print** — the default for new songs, as lyrics have always been.
  - **All together — at the start / at the end** — every section's words
    in one block, in song order, each under its section's name.
  - **All together — left column** — the same block down a column of its
    own on the left, with the chart in one column beside it on every page.
    Columns steps aside (and says so); words longer than the chart get
    pages of their own, which "fit to one page" shrinks the type to avoid.
  - **Each section — beside chart** — v0.23's layout: the words to the
    right of their own chart, showing what you play during them.
  - **Each section — under chart** — the words under their own chart.

  The gathered modes use each section's words. If no section has any yet
  they fall back to the whole-song sheet, with its `=== Verse 1 ===`
  markers promoted to headings — the one place a marker earns ink. A
  per-section layout on a song whose sheet was never split does the same
  at the start, so words the song has are never silently left off.
- **⚙ Layout, bottom left.** Page size, columns, colour, section names,
  lyrics, recalled licks and chord shapes, in one panel grouped by what
  they affect, each with a line saying what it does. A summary under the
  button ("Fit to page · Auto · lyrics: beside") shows what's set without
  opening anything. It's a popover, not a modal: the song stays visible
  and the preview keeps updating while you change things. The desktop app
  has the same settings behind a ⚙ Layout button on a slim bar along the
  bottom of its window.

### Changed
- **The Print & Export strip is gone from the song form**, which is back
  to being about the song: title, artist, key, time, BPM.
- **No more per-section "print" checkbox** in the Lyrics dialog or the
  blank-line split. It now says where the words are printing, with a
  ⚙ Layout… button that jumps to the setting. The dialog is for the
  words; Layout is for the page.
- **Songs saved before this version keep printing what they printed.** A
  song with no placement set opens as "each section — under chart" if any
  section's box was ticked, "at the start" if only the whole-song sheet
  was, and "don't print" otherwise.
- Shorter option labels throughout ("Tab again", "At the end", …) now that
  each sits next to a label saying what it's for.

### Fixed
- **"Chord shapes first" crashed the PDF export** — the sheet was drawn
  before its drawing code was defined. Only "at the end" had a test.

### API
- `render.py`: `lyrics_mode()`, `section_lyric_lines()`, `lyric_blocks()`,
  `gathered_lyrics()`, `lyric_block_items()`, `lyric_block_lines()`;
  `LYRICS_MODES`, `LYRICS_GATHERED`, `LYRICS_PER_SECTION`.
- `model.py`: `LYRICS_LAYOUTS` now has six values; `migrate_document()`
  fills in `lyrics_layout` for older songs.
- `export.resolve_columns()` returns 2 for the left-column layout,
  whatever `pdf_columns` says.
- Sections' `print_lyrics` flags no longer affect printing. They're still
  written when lyrics are assigned, so an older version opening the file
  makes a similar call.

### Notes
- 319 tests pass (was 305).

## [0.23.0] — 2026-09-20

The sheet marks itself up, the words move out of the chart's way, a lick
you've played once gets a name, and the chord shapes you had to work out
print on the sheet instead of in your head.

### Added
- **Section markers in the lyric sheet.** Write `=== Verse 1 ===` above
  the words that belong to it — or drag the section's chip into the sheet
  from the Lyrics dialog, which drops the marker for you — and **Split by
  markers** hands each block to the section it names. Marker names match
  loosely (`Chorus_1` finds "Chorus 1"), a section the sheet never names
  is left exactly as it was, and a section marked twice keeps both halves.
  Names with no section behind them are offered as sections to create,
  which is also what **+ Section…** in that dialog does from scratch.

  The old blank-line split is still there as **Split on blank lines…**,
  for a sheet pasted straight off a lyrics site. It asked you to hold the
  shape of the song in your head while reading a list of first lines;
  markers say the same thing where you can see it, survive re-editing the
  words, and leave the sheet self-describing next time you open it.

  Markers are structure, not words: they're drawn in colour in the editor
  (a tag on the Text widget in the desktop app, an overlay behind the
  textarea in the browser) and stripped from everything that prints.
- **Named licks.** `{Riff1 = G 5 7 5 | D - - 3}` names a lick; `{Riff1}`
  recalls it anywhere else in the song, with a repeat if you want one:
  `{Riff1}x3`. It resolves at render time to the notes themselves, with
  the name printed over them — which is how the handwritten charts say it
  ("RIFF 1 (x3)"), and it means editing the lick once updates every place
  that plays it, exactly as a section reference does.

  There is no separate lick library to keep in step: a name lives on the
  lick where it was written, and `songmap.lick_index()` derives the index
  by walking the song. Both front ends fill the name in for you — the
  `{ name = tab }` button inserts `Riff1`, `Riff2`, … — and the browser's
  palette offers every lick already named as a button that inserts the
  reference.

  What a recalled lick prints as is a setting, `lick_refs`: *tab* (the
  default) plays the notes again wherever it's used; *name only* prints
  `Riff1 (x3)` and leaves the notes where the lick was written — the
  shorter page, and how the handwritten charts do it. Either way a lick's
  name prints in the lick blue, and a reference that finds nothing keeps
  its braces (`{Nope}`), so a typo is visible on paper.
- **Chord shapes, printed once.** A new **Chords 🎸** dialog keeps a list
  of voicings written the way a chord chart writes them (`x32010` is C,
  `x 0 12 12 12 x` when the frets need two digits), and prints them as a
  block of tab-style diagrams at the start or the end of the chart —
  `chord_sheet` in the `.sng`, `"none"` by default and set to `"end"` the
  first time you add a shape. Shapes are grouped by instrument, laid out
  in as many diagrams per row as the page width takes, and each can carry
  a word of its own ("barre", "thumb"). New module: `chords.py`.

  The dialog checks the shapes against the chart: it says which chords
  are played with no shape yet and which shapes are for chords the song
  never plays, and **+ From chart** adds a row for every missing one.
  "Played" means *as printed* — references expanded, transpose applied —
  so a song moved up a tone asks for D, not the C you typed. Guitar
  sections only when the song has any, since on a bass chart "A" is a
  root, not a chord to finger. `cli.py lint` reports the same two lists
  as notes, never as failures.

  **Shapes move with the song.** Transpose the song and the chord sheet
  follows — re-voiced the way a player would, not by sliding every fret:
  1. a shape with no open strings (a barre, a power chord) is already
     movable, so it **slides** — same hand, different fret;
  2. an open shape goes to the **open shape** of the new chord if there
     is one — C up a tone is `xx0232`, not `x54232`;
  3. if there isn't (no open F, no open Bm) it goes to the **E-form or
     A-form barre**, whichever sits lower on the neck;
  4. a chord with no barre template (an add9, a slash chord) is slid
     whole, open strings included — the capo answer.
  Off standard six-string tuning only 1 and 4 apply. As with everything
  else it's render-time: the shape you typed is what's stored. The dialog
  says what each row prints as ("→ D · open shape") while the song is
  transposed, coverage compares shapes *as printed*, and **+ From chart**
  stores each new row under its un-transposed name so it prints as the
  chord the chart shows.

### Changed
- **A section's words print beside its chart, not under it.** The chart
  keeps a narrow left column and the words run down their own column to
  the right of it, starting level with the chart's first line. Nothing is
  aligned chord-to-syllable — that would be a claim about where the
  changes fall that a chart like this can't honestly make. What it says
  is *during these words, this is what you play*, which is the reason to
  print them together at all.

  The practical difference is vertical: a verse now costs the page the
  taller of the two sides rather than the sum of them, so the section
  after it doesn't get pushed off the sheet. `lyrics_layout` in the
  `.sng` — `"beside"` by default, `"below"` for the old behaviour — is in
  the browser's Print & export row and in the desktop Lyrics dialog. The
  width bound the PDF fit works against and the never-split-a-section
  height estimate both see the new layout, so the page still measures
  itself correctly.

### API
- `lyrics.py`: `split_marked()`, `match_sections()`, `apply_marked()`,
  `unmatched_names()`, `strip_markers()`, `marker_line()`, `is_marker()`,
  `has_markers()`.
- `render.py`: `printable_lyrics()`, `lyrics_beside()`,
  `lyric_column_x()`, `compose_beside()`, `beside_line_count()`.
- `chords.py`: `shape_from_text()`, `shape_to_text()`, `diagram_lines()`,
  `sheet_lines()`, `sheet_position()`, `strings_for()`, `used_symbols()`,
  `coverage()`, `transpose_chord()`, `printed_chords()`, `stored_name()`,
  `split_chord_name()`, `is_open_shape()`.
- `songmap.py`: `lick_index()`, `find_lick()`, `lick_names()`.
- `model.py`: `make_lick(lines, name, repeat)`, `make_lick_ref()`,
  `make_chord()`; new document keys `lyrics_layout`, `lick_refs`,
  `chords`, `chord_sheet`; new item kind `lick_ref`.
- `webserver.py`: `POST /api/lyrics/marked`, `POST /api/chords/shape`,
  `POST /api/chords/coverage`; `/api/meta` gains `lyrics_layouts`,
  `chord_sheets` and `lick_refs`.

### Notes
- `{G}` used to be a parse error (a lick line needs frets as well as a
  string name); it is now a reference to a lick named `G`. Nothing that
  parsed before parses differently.
- 305 tests pass (was 241), including a new `tests/test_chords.py`.

## [0.22.0] — 2026-09-20

Two columns. A chart is mostly short lines — four chord symbols, a bar
of tab — so one column down an A4 page leaves half the sheet white, and
**Fit to page** can only grow the type until that single column is full.
Splitting the page halves the height the chart needs, and the fit spends
that height on type size instead: the same song, printed bigger, and
often on one sheet instead of two.

### Added
- **Columns — a new Print & export setting.** *Auto* (the default), *One
  column* or *Two columns*, stored in the `.sng` as `pdf_columns`
  alongside `section_layout`, `color_mode` and `pdf_scale`, so a chart
  prints the same way wherever it's opened. It's in the meta panel in the
  browser and in the Export PDF dialog in the desktop app.

  A two-column page is filled top to bottom down the left column and then
  down the right, with a grey rule drawn down the gutter — a page split
  you can't see is two charts you have to guess the edges of, which is
  the opposite of useful on a stand mid-song.
- **`export.resolve_columns()` and `export.column_width()`.** The column
  count is resolved from the document the way the scale already was, and
  every measurement that used to ask for a page width — the width bound
  the fit works against, the tab grid's measures-per-line — now asks for
  a *column* width, so one column and two run through the same code.
  `export.build_pdf()` is unchanged for callers: the CLI, the web server
  and the desktop app all pick the layout up for free.

### Changed
- **Fit and columns are decided together.** Half a page's width fits a
  different size of type than a whole one, so `resolve_scale()` takes the
  column count (defaulting to whatever the document asks for), and *Auto*
  builds both layouts and compares them on what matters at a music
  stand — in this order: how many sheets you have to turn, then how big
  the type is. A page saved beats a point of type size; reaching for the
  next sheet mid-song costs more than a slightly smaller chord symbol.
- **Auto knows when not to split.** A chart of long lines comes out
  *smaller* across two columns, not bigger, and a tab grid can wrap but
  not below one measure per line. Either case leaves the page whole
  rather than printing something that overflows its column.
- **Bigger notes.** The printed body goes from 7.5pt Courier to 9pt (and
  the section headings from 9 to 10.5), a fifth larger at the same fit.
  Courier sets small for its point size, and a chart is read at arm's
  length off a stand — the old body read noticeably smaller than the
  headings beside it. Size and character width move together
  (`export.MONO_SIZE`, `export.MONO_CHAR_W`: Courier's advance is exactly
  0.6 em, and that advance *is* the column arithmetic), so nothing drifts
  out of its column; the line height is unchanged, so this costs no page
  count.
- **A tab block is spaced the same on both sides.** A lick with no chord
  symbols beside it already had a blank line above it; the chart line
  after it butted straight up against the tab, so the block read as
  open-ended. It now closes with the same blank line — except at the end
  of a section, where the gap is already there.
- **Fret numbers are set as a figure over the chord.** Smaller than the
  symbols (`constants.FRET_SIZE_RATIO`), raised close to the line they
  belong to rather than taking a full line of their own
  (`FRET_LINE_RATIO`), and in amber on colour or a light grey in black &
  white (`FRET_RGB` / `FRET_BW_RGB`) — a fret is a hint about *where* to
  play the chord under it, not a second line of notes. Same treatment in
  the browser's section preview.
- **The printed chart is drawn on its columns.** The renderer aligns a
  chart by padding with spaces and counting characters; the page draws in
  a proportional face, where a space is nothing like a character wide, so
  a row drawn as one string landed near its columns and drifted further
  along the line. Each stretch of ink is now placed at the column it was
  rendered at — the padding is measured, not drawn. Most visible on the
  thing that has to sit over its chord: the fret number.
- Columns are a PDF concern: the TXT export is unchanged. The tab spacing
  above is a rendering fix, so it shows in TXT, the Preview pane and
  Stage View too.

- **Split a lyric sheet across the sections, and reuse the blocks.**
  Paste (or import) the whole song's words once, then **Split into
  sections…**: the sheet is cut on its blank lines and each section gets a
  row with a dropdown of the blocks, so the same chorus goes to Chorus 1,
  2 and 3 in one pass. The proposal is made by the new `lyrics.py` and
  shared by both front ends — a repeated block is the chorus and goes to
  every Chorus and Refrain section; instrumental sections (Intro, Solo,
  Interlude, Breakdown, Outro) are stepped over rather than being handed
  verse one.

  The old split assigned block N to section N, which put the first verse
  on the intro of any song that opens with one, knocked every later block
  one section out of place, and reached only the first of three choruses.

  The whole-song sheet is now kept rather than cleared, so a section added
  next week can still be given one of its blocks — it just stops printing,
  so the same words don't land on the chart twice. A section holding words
  that came from somewhere else can be left alone ("keep what's here").
- **The Lyrics dialog lists the song's sections.** A row of chips under
  the Scope dropdown, one per section, each with a dot that's filled when
  that section already has words — the gaps are visible without opening
  every section in turn, and getting to one is a click. The dropdown
  itself now names each section's type and marks the ones that have words.

### Fixed
- **A fret number went black after the first `//` block in the browser.**
  The server told the front end "the first row is frets" by asking
  whether there were fret numbers *anywhere* on the line. A line whose
  opening block has none starts on symbols, so the browser painted a
  symbol row as frets and the real fret row as symbols. It now reports
  the first row's own role, and every fret row is styled by role wherever
  it appears.

## [0.21.0] — 2026-09-20

A lick is tab, so it now looks like tab. Three small changes that all
come from the same complaint: on a printed chart, a `{G 5 7 5}` sitting
in a line of chord symbols reads as more symbols until you look twice —
which is exactly the moment you don't have, mid-song, glancing down at a
music stand.

### Added
- **`{ tab }` — an empty lick, built for the instrument you're on.** The
  quick-insert palette's lick button no longer drops a worked example to
  be deleted; it inserts a blank grid with every string of *that
  section's* instrument already named and six positions of dashes ready
  to type frets over:

  ```
  {G - - - - - - | D - - - - - - | A - - - - - - | E - - - - - -}
  ```

  The caret lands on the first position. Six is a starting width, not a
  limit — add or remove dashes and the lick is as long as you write it.
  The button is in both front ends: beside the chart line in the browser,
  and beside the editor bar in the desktop app, where it reads the
  focused section's (or riff's) instrument the same way.
- **Row roles, so a front end can colour what it draws.**
  `render.render_chart_rows()` and `render.chart_body_rows()` return each
  rendered row as `{"text", "role", "spans"}` — the role being `fret`,
  `sym`, `lick` or `text`, and `spans` marking stretches *within* a row
  that print differently. The plain-text `render_chart_row()` and
  `chart_body_lines()` are unchanged and now wrap these, so the TXT
  exporter and every existing caller carry on as before. `/api/parse`
  returns `roles` and `spans` alongside `rendered`.

### Changed
- **A lick prints as tab**, one row per string, the string named in a box
  at the left, instead of as a bare row of numbers:

  ```
  |G|-5-7-5-|
  |D|-----3-|
  ```

  Every position is the same width down the whole lick, so the columns
  line up vertically and it reads the way a tab staff reads. A
  two-digit fret widens every cell in that lick rather than just its own
  — misaligned columns would defeat the point of stacking the strings.
- **A lick prints blue**, on screen and in a colour PDF. Not decoration:
  it's the one thing on a chart line that isn't a chord symbol, and the
  colour is what makes it findable at a glance. Black-and-white mode
  prints it dark grey rather than dropping the distinction, since tone is
  all a mono printer has.
- **A rest prints grey.** It's the absence of playing; it shouldn't
  compete for attention with the notes on either side of it. Drawn as its
  own coloured run inside the symbol row, so the rest of the row stays
  black.

## [0.20.0] — 2026-09-18

Two ways to stop fighting the notation, and one launch annoyance
fixed. The chart-line grammar has grown enough that remembering
whether a second ending is `|2.` or `[2]` costs more than typing the
note does — so the syntax is now a row of buttons beside the line it
edits. And for the bits the grammar genuinely doesn't cover, a section
can now just be text: `render: "free"` turns the whole chart apparatus
off and prints what you typed. That's a deliberate escape hatch, not a
gap to close later — a one-page stage chart is allowed to have a
sentence on it.

### Added
- **Fit to one page** — a new Size option beside Fit to page. Fit to page
  keeps the page count the chart already had and grows the type into
  whatever room is left; this one fixes the page count at one and shrinks
  the type until the whole song lands on a single sheet, so there's no
  page turn mid-song. It stops at a readability floor (30%) rather than
  shrinking without limit: a chart that still needs two pages there is
  printed at the floor and honestly runs long.
- **Quick-insert palette** in the web UI — a compact column of buttons
  beside each section's chart line: `rest`, `%`, `|:`, `:|`, `|1.`,
  `|2.`, `x2`, `[ ]x2`, `""`, `segno`, `coda`, `dc`, `ds`. Each drops
  its notation at the caret, space-separated from what's already
  there, and leaves the caret *inside* the brackets or quotes where
  that's the useful place for it. The insert fires the same `input`
  event typing would, so the existing debounced parse, inline error
  display, and live preview all run unchanged — the chart line stays
  typed text you can still edit by hand. The palette moves below the
  line on narrow windows rather than squeezing the input it serves.
- **Break a section over several lines.** `//` anywhere on a chart line
  starts a new line there, so a long section can be grouped into the
  phrases you'd write out by hand instead of running as one wide row:

  ```
  A(5) D(5) // F(8) D(5) E(7) // A(5) A(5) F(8)
  ```

  Each line aligns its own columns — that's the point of breaking, rather
  than one grid stretching across the whole section — and in the desktop
  song map the continuation lines stack under the section name rather
  than sliding back to the margin. Purely a layout mark: it plays
  nothing, the chart line stays a single editable line, and the TXT, PDF,
  Print and Preview all honour it. `//` is on the quick-insert palette.

  Adding `>` pushes that line in — `//>` for one step, `//>>` for two — so
  a phrase can sit visibly inside the one above it:

  ```
  A(5) D(5) //> F(8) D(5) E(7) // A(5) A(5) F(8)
  ```
  ```
    A(5)  D(5)
          F(8)  D(5)  E(7)
    A(5)  A(5)  F(8)
  ```

  The indent is relative to whatever the line would otherwise start at,
  so it behaves the same in the export and in the desktop song map, where
  lines already sit in the section-name gutter. Stored as an `indent`
  count on the break mark, and left off entirely when it's zero, so
  existing documents are byte-identical.
- **Licks — a bit of tab, in among the chords.** Some things aren't a
  chord. `{G 5 7 5 | D - - 3}` in a chart line writes a short tab figure
  at the point it's played: one `|`-separated line per string, its name
  then its frets, positions aligning by column across the lines, `-` for
  not played and `x` for muted. It prints as a small tab block under its
  own column in the chart row, so it stays in sequence with the chords
  rather than being exiled to a separate grid — as many string lines as
  the figure needs. Transposing a section moves a lick's frets with it
  (same strings, shifted positions), and a fret that would fall off
  either end of the neck is left alone rather than silently clamped to a
  position you'd actually play. `{ }` is on the quick-insert palette.
- **Free-text sections** — a fourth per-section render mode, `free`,
  alongside Chart / Tab / Both. The section shows one plain textarea;
  nothing in it is parsed, transposed, validated, or column-aligned,
  and the TXT and PDF exports print it verbatim. Stored as a new
  `free_text` string on the section. Switching a section to Free
  leaves its existing items untouched, so switching back to Chart
  brings the chart line back exactly as it was — the two are
  alternative views of one section, not a destructive conversion.
  `render.estimate_section_lines()` counts free text toward the live
  page estimate, so the page count still tracks the real export.
- **Native window without any setup** — `webserver.py` now probes for
  a sibling virtualenv (`.venv`, `venv`, or `env`) whose interpreter
  can `import webview`, and re-execs into it (`os.execv` — same PID,
  same argv, no wrapper script). So `python3 webserver.py` with the
  system interpreter opens the native app window instead of silently
  falling back to a browser tab just because the system Python can't
  see `./.venv/lib/.../pywebview`. `SNT_NO_REEXEC=1` disables it, the
  child always carries that flag so a broken venv can't produce an
  exec loop, and `--browser` / `--no-open` skip the probe entirely.
  The "am I already inside that venv?" test compares `sys.prefix` to the
  venv directory, **not** the interpreter paths: `.venv/bin/python3` is
  normally a symlink straight back to the base interpreter that created
  the venv, so on a machine where `python3` *is* that base interpreter
  (a python.org framework build on macOS, most distro Pythons on Linux)
  comparing `realpath(sys.executable)` matches even though the running
  process has none of the venv's packages — which is precisely the case
  the re-exec exists to fix. If a venv is found but its interpreter
  can't import `webview`, that now prints one line saying so with the
  underlying error, instead of silently opening a browser tab;
  `SNT_DEBUG_LAUNCH=1` prints the full traceback.

### Changed
- **The wiki is published from this repository.** `wiki/` is now the
  source of truth and a workflow copies it to the GitHub wiki when a
  change lands on `main` — so a documentation edit is reviewed in the
  same commit as the code it describes, and one push updates both.
  Previously the wiki was a separate clone with its own commits, which
  meant it could sit at a different revision from the code indefinitely
  (and did). The sync is one-way and total, so `wiki/` is the only place
  to edit; the workflow can be re-run by hand from the Actions tab if the
  wiki ever drifts.
- **Every export carries the project link in its footer** — TXT and PDF
  alike. A chart handed to a bandmate is the only place its reader can
  look to find out what made it. The URL lives in `constants.APP_URL`,
  beside the version, so it's stated once.
- **Section names print without square brackets** in both TXT and PDF.
  The brackets marked the name out from its surroundings, but a banner
  band or a divider rule already does that — they were two more
  characters between the reader and the name.
- The browser fallback message now gives the two commands that
  actually fix it (create a venv, install `requirements-optional.txt`)
  instead of a bare `pip install pywebview` that lands in whichever
  interpreter happens to own `pip`.
- The desktop app renders a free-text section read-only in the song
  map (first six lines, then an ellipsis) and disables its editor bar
  with a hint pointing at the web UI, rather than showing an empty
  chart row for a section that has no chart items.
- "Duplicate as reference" on a free-text section produces a normal
  chart section — a `section_ref` has no free text of its own to show.
- **The Notation Reference now explains the marks rather than naming
  them.** Every mark gets its real name, where the word comes from, and
  what it actually tells a player to do — `segno` as the place-marker
  `ds` jumps back to, `coda` as the separate tail you break off to, the
  1st/2nd endings as the "two different last bars" case — plus a note
  on the standard D.S.–segno–coda shape and why it's three words
  instead of three near-identical pages. A chart nobody but its author
  can read isn't much of a chart.

- **Print, through the OS's own print dialog.** A "Print…" item in the
  Export menu builds the PDF and hands it to the system viewer, so you
  get the real print panel — printer, paper size, scaling, page range —
  and a preview before anything reaches paper. Deliberately not a silent
  `lp` job: a stage chart is exactly the kind of thing worth eyeballing
  first. In a browser tab it opens in a new tab instead, where Cmd/Ctrl+P
  does the same.
- **Colour, or black and white.** The PDF now colours each section's
  *heading* by section type, so every Verse looks like every other Verse
  and you can find your place on a stand at a glance. The notes
  themselves stay black — colour is a label, and tinting the notes only
  makes them harder to read. The palette mirrors the accent each section
  card carries in the editor, darkened for white paper (a colour that
  reads well on a dark background is washed out in print). **Black &
  white** fills every heading band with a light grey and sets the text in
  black, for a mono printer or a photocopy where a palette just becomes
  indistinguishable greys — grey rather than solid black because a
  full-width black band per section is a lot of toner for something that
  only has to say "new section starts here", and black text on grey
  survives a photocopy better than white reversed out of black.
- **One instrument, stated once.** When every printed section uses the
  same instrument, it moves up beside the key and tempo in the header
  instead of being repeated on every section heading — it's a fact about
  the song, not about each section. Songs that genuinely switch
  instrument keep the per-section label. Filtering the export to one
  instrument counts as a single-instrument chart and reads like one.
- **Fit to page.** A new Size setting scales the whole chart up until it
  fills the sheet without needing another page — a chart you can read
  from a music stand, or off the floor. It's exact rather than a guess:
  the width bound is computed from the longest line, and the height is
  bisected on real builds, with "fits" meaning "needs no more pages than
  it did at 100%". It shrinks too — a chart whose longest line already
  ran off the edge of the paper was never "fitted" by leaving it there —
  down to a legibility floor, past which landscape is the honest answer.
  Fixed percentages (100/125/150/200%) are there for when you want a
  specific size regardless.
- **Sections on the left, optionally.** A new per-song "Layout" setting
  puts each section's name in a left-hand column beside its first line
  instead of a full-width header band above it. That's three lines saved
  per section — for most songs the difference between a one-page chart
  and a two-page one, which is the whole point of this program. Stored in
  the `.sng` as `section_layout`, so a chart prints the same way wherever
  it's opened, and honoured by TXT, PDF, Print and the Preview pane
  alike. Defaults to the existing banner style; nothing changes for an
  existing song.
- **Your files live in your home directory, and the app says where.**
  Songs were written to `./songs` inside the checkout. They were
  gitignored, so they never got committed — but "not in the repo's
  history" isn't "safe": a `git clean -xdf`, a fresh clone, or deleting
  the checkout took the whole folder with it. Songs are the user's data,
  not part of this program, so they now default to
  `~/Documents/Song Notation Tool`, where the usual backups cover them.
  Anything still in the old folder is moved there on first run and
  reported by name; a name that already exists in the target is left
  alone rather than overwritten, and the old folder is never deleted.
  The folder is shown in the sidebar with **Reveal** and **Change…**
  buttons (native window only — a browser tab can't open a folder
  picker), and the choice is remembered in a small config file under
  `~/Library/Application Support/Song Notation Tool/`. `--dir` still
  overrides it for one run without changing the setting.
- **Export asks where to put the file.** In the native window an export
  now opens a real save panel, pre-filled with `<Artist> - <Title>.pdf`
  and starting in the folder you exported to last, then reports the full
  path it wrote. Previously the export was an HTTP download with no
  download UI behind it in the native window — which is a good way to
  lose a file. In a browser tab it's still an ordinary download (there's
  no save panel to reach from there), but it now says so.
- **References print what they play.** A `section_ref` / `block_ref` is
  stored as a pointer — that's what makes "edit the riff once, every
  section using it updates" work — but the chart used to render the
  pointer itself, so a duplicated-as-reference section printed `=chorus1`
  where the notes should be. The pointer stays in the data and is now
  expanded at render time (`render.resolve_references`), in the section
  card's preview row, the Preview pane, and the TXT and PDF exports. A
  repeat expands into a bracketed group (`[3A 12D](x2)`), a `+N`/`-N`
  shift expands transposed, and a reference that can't be resolved — a
  missing target, or a cycle — is left as a reference rather than
  silently dropping the section's content.
- **References are now *displayed* by section name, always.** Writing
  them by name (below) only helped new references; an existing
  `=chorus1` in a saved song still read as a stale id. The chart line is
  now a view: a reference is spelled with its target's current name
  wherever it's shown, and whatever you type — `=Interlude`, `=chorus1`
  — is stored as the target's id. So the display stays readable, the
  document stays rename-safe, and renaming a section updates every
  reference to it on screen without touching a single stored reference.
  A name that can't be written as a reference (punctuation, or two
  sections sharing it) still shows the id, which is never ambiguous;
  spaces and dashes display as underscores (`=Verse_2`) and resolve back.
- **References can be written by section name, not just id.** Ids are
  minted once and never change, so a section created as "Chorus" and
  later renamed "Interlude" keeps id `chorus1` — correct, but `=chorus1`
  on screen reads like a mistake. `=Interlude` now resolves too
  (case-insensitive, spaces and dashes interchangeable with
  underscores), ids still win on a tie, and "Duplicate as reference"
  writes the name whenever it's a bare identifier and unambiguous.
  Existing documents are unaffected — every id-based reference still
  resolves exactly as before.

### Fixed
- **A section could still be split across a page break.** The rule that
  keeps a section whole measured it with `estimate_section_lines()`, which
  didn't expand references — so a section whose entire content is
  `=Chorus_1` measured one line, was judged to fit in the space left at
  the foot of the page, and then printed seven lines across the break.
  The estimate now renders through references when it's given the
  document, and allows for the heading band and the inter-section gap.
- **A hidden tab grid printed anyway.** The editor shows the tab grid only
  in Tab or Both mode, but the exporter drew any stored measures whatever
  the mode — so a section switched back to Chart printed an `M1` tab block
  that nothing on its card accounted for. Both exports now follow the
  render mode: Chart prints the chart line, Tab prints the grid, Both
  prints both. (Switch such a section to Both to see, and delete, the
  measures it's still carrying.)
- **A referencing section's card went stale when its target changed.**
  A `=Chorus_1` card shows the target's content, but nothing in the page
  recorded that dependency: adding a `//` to Chorus_1 re-rendered Chorus_1
  and left Chorus_2 showing the old, unbroken row until the file was
  reloaded. The exports were always right; only the on-screen card lagged.
  Editing a section's line, free text, render mode or name now refreshes
  every card that holds a reference.
- **Renaming a target left the reference pointing at nothing on screen.**
  The card re-parsed its typed `=Old_Name`, which after the rename matched
  no section, so the preview fell back to printing the reference itself.
  A referencing card is now rendered from its *stored* items — which hold
  the target's id — via a new optional `items` field on `/api/parse`, so
  the preview stays correct and the line respells itself with the new name.
- **Curly quotes were a parse error.** macOS — and every word processor —
  substitutes a typed `"` with `"` or `"`, so an annotation typed
  anywhere but a plain terminal failed with `cannot parse item`, pointing
  at the text rather than at the character. All four quote characters are
  now accepted.
- **A long title ran off the page at high scale.** Fit-to-page bounds the
  width using the monospace body, but the title is set in a proportional
  face and isn't covered by that measurement — at 2.5x a long title
  simply overflowed the right margin. The title and meta line now get
  their own size cap and shrink on their own rather than dragging the
  whole chart's scale down with them.
- **A section card's preview only ever showed two lines.** Fine while a
  chart row was always a fret row and a symbol row; a lick adds one line
  per string. The row is now rendered in full, and `/api/parse` reports
  whether the first line is frets so the front end doesn't have to guess
  from the count (it was guessing wrong for a row with no fret numbers,
  colouring the chords as frets).
- **Sections were squeezed against their own headings.** The gap between
  a section's name and its first line, and between one section and the
  next, was barely a line — readable at 100%, cramped at any larger
  scale, and hard to scan on a stand. Both gaps are now generous and
  scale with the type.
- **The print settings looked like song data.** Layout, Colour and Size
  sat in the same row as Title, Artist, Key, Time and BPM, so they read
  as five more facts about the song rather than three choices about how
  it comes out on paper. They're now on their own second line, in a
  dashed, tinted frame labelled "Print & export" — nothing inside it
  changes a single note.
- **The Layout control didn't match the fields beside it.** A `<select>`
  needs both the shared styling *and* `appearance: none`, or macOS draws
  its own white pill over the top — which is what made it look pasted in
  from another application. All three meta selects now match the inputs
  exactly, chevron included.
- **The toolbar and song list scrolled away.** The page was an ordinary
  scrolling document, so working on a section near the end of a long song
  left Save, Preview, Export and the song list all off-screen — you had
  to scroll back to the top to do anything with what you'd just typed.
  The window is now an app shell: a fixed-height column where only the
  section list scrolls, with the toolbar, the song list and the songs
  folder always in view. The Preview pane and the help strip scroll
  within themselves.
- **A reference to a free-text section showed the wrong content.**
  Switching a section to Free deliberately keeps its old chart items, so
  switching back is lossless — but expanding a reference walked those
  dormant items and printed a chart the target itself no longer displays.
  A reference now renders whatever its target renders: free text for a
  free section, the chart row for a chart one.
- **A chart row with no fret numbers printed a blank line above itself.**
  `render_chart_row` always emitted both rows, and an all-blank fret row
  `rstrip()`s to an empty string — so any section of plain chords, or a
  lone expanded reference, sat one line lower than its neighbours.
- **A section's annotation could be set but never removed.** Typing
  `"like this"` on a chart line stored it as the section's `annotation`,
  `grammar.unparse()` then (correctly) left it out of the line, and
  nothing anywhere displayed it — so once set, a quoted note was
  permanent, and in a free-text section, where the chart line is hidden
  entirely, it couldn't even be re-typed. Each section card now has a
  visible Annotation field, in every render mode; clearing it removes the
  annotation. Typing `"…"` on the chart line still works and now lands in
  that field instead of disappearing into the document.
- **Chart sections printed two columns right of everything else.**
  `render_chart_row()` used a minimum 4-column gutter even with no label,
  while section headers, annotations and free text all indent by 2 — so a
  chart or referenced section sat visibly out of line with the free-text
  sections around it. There's now one `render.BODY_INDENT` the export
  uses for every body line. A labelled row (the desktop song map's name
  gutter) is unchanged.
- **A group's symbol row dropped its fret numbers** — `[3A 12D]x2`
  printed as `[A D](x2)`. Harmless while groups were only ever typed by
  hand; not harmless once an expanded reference renders as one.
- **Section cards wasted most of the window.** Three separate causes,
  found by measuring the real layout in a headless browser rather than
  eyeballing it:
  - `#main` was capped at `max-width: 980px`, so a 2000px-wide window
    left ~800px of dead space to the right of every card. It now fills
    the width it's given, with a 1680px upper bound (and auto margins
    past that, so an ultra-wide display doesn't stretch one line of
    chords across half a metre).
  - `.sec-preview` carried `white-space: pre`, which preserved the
    newlines and indentation of the template's own markup between its
    two child rows — an *empty* preview still rendered about six blank
    line boxes, ~117px of dead height in every card. `pre` now sits on
    the two rows, where the column alignment actually needs it.
  - The quick-insert palette was pinned to a fixed 148px column, so it
    always wrapped to four button-rows and made every card as tall as
    the palette rather than as tall as the input. It's now sized by its
    content: one row beside the input in a wide window, wrapping only as
    the window narrows. Net effect on a wide window: a section card went
    from 294px tall to 117px.
- **The page scrolled sideways below ~900px.** `.section-head` packs
  nine controls into a non-wrapping flex row, so the whole document grew
  wider than the viewport instead of the header wrapping. Both it and
  the topbar now wrap. Also fixed the narrow-window rule flipping
  `.sec-line-row` to `flex-direction: column`, where the input's
  `flex-basis` sizes *height* — it was rendering 260px tall.
- **A section's rendered chart row was blank until you typed in it.**
  `updateSectionPreview()` treats "no parse result" as "keep the last
  render", but a freshly built card has no last render to keep, so
  every card from a saved song opened with an empty preview. Cards now
  fetch their render once on build.
- **Icon buttons in the section header had their glyphs off-centre**
  (the ↑ ↓ ⧉ 🗑 row, and the sidebar's +). They carry both `.btn` and
  `.icon-btn`; `.btn`'s `padding: 6px 12px` was being applied inside
  `.icon-btn`'s fixed 28×28 box, leaving almost no content area and
  pushing the glyph to one side. Now centred with flex and zero
  padding — the glyphs in use have very different bearings and
  baselines, so `text-align`/`line-height` alone can't place them
  reliably.

## [0.19.0] — 2026-09-17

Two roadmap items land together: the browser front end gets the same
non-destructive Transpose the desktop app has had since v0.4, and a new
Lyrics panel — reference text alongside a section or the whole song,
in both front ends.

### Added
- **Transpose in the web UI** — a "Transpose ↕" button in the browser
  toolbar opens the same whole-song-or-one-section choice as the
  desktop app's Transpose dialog, by any number of semitones. It sets
  `doc.transpose` / `section.transpose`, the exact fields the
  render-time engine (`transpose.py`) and the desktop app already use,
  so a song transposed in the browser looks right in the desktop app
  and vice versa — nothing device-specific about it.
- **Lyrics panel**, desktop and web — a new "Lyrics" button opens a
  scoped text box (whole song, or the focused/selected section) with
  three ways to fill it: type or paste directly, **"Import from
  file…"** to load a local `.txt` file, or **"Search lyrics online
  ↗"**, which asks for the song title and artist (pre-filled from the
  doc's meta fields, editable — they might be empty or describe this
  arrangement rather than the published song) and opens a DuckDuckGo
  search for `<artist> <title> lyrics` in a new tab — it never fetches
  or auto-inserts anything, by design (both copyright and
  section-boundary accuracy: you decide what to paste and where it
  goes). Stored non-destructively as `lyrics_text` on the document or
  the section, alongside `transpose` — it's a reference layer, never
  parsed or aligned to the chart, and **off by default** in the
  TXT/PDF export and the Preview pane.
- **"Include in TXT/PDF export and the Preview pane"** checkbox in the
  Lyrics panel (desktop and web) — sets a new `print_lyrics` flag on
  the document or section, independent of `lyrics_text` itself, so
  pasting reference lyrics in never changes what a song prints until
  you opt in. `render.py`'s page-count estimate and the PDF's
  never-split-a-section-across-a-page-break guarantee both account for
  the extra lines when it's on.
- **"Split into sections…"** in the Lyrics panel (whole-song scope
  only, desktop and web) — splits the pasted-in text on blank lines
  and assigns one block per section in document order: existing
  sections first, then a new section per leftover block. Turns "paste
  the whole song's lyrics from a search result" into one step instead
  of copying each verse/chorus in by hand via the scope picker; asks
  for confirmation first since it overwrites each target section's
  existing lyrics text.
- **`webserver.py` exposes a small `js_api`** (`open_url`) to the
  front end when running as a native `pywebview` window — that
  window has no browser-tab concept, so `window.open()`/`target=
  "_blank"` is a silent no-op there; "Search lyrics online" now goes
  through `window.pywebview.api.open_url()` (falling back to
  `window.open()` in plain browser mode) so it reaches the OS's
  actual default browser either way.
- `model.py`: `new_document()` and `new_section()` now default
  `lyrics_text` to `""` and `print_lyrics` to `False`. Existing `.sng`
  files (including ones migrated from v0.15) load fine without either
  key — every read goes through `.get(..., default)`.
- Help strip (web) and in-app help (desktop, "?" button) updated to
  describe both features.

## [0.18.1] — 2026-09-17

A follow-up to `webserver.py`'s launch mode from v0.18.0: a resizable
native app window instead of a browser tab, and a proper in-app "Save
& Close" alongside it.

### Added
- **`webserver.py` opens a native app window**, when `pywebview` is
  installed: a normal titled OS window (resizable, its own
  minimize/maximize/close buttons) rather than a full browser tab with
  an address bar and tabs — the same idea as the sibling
  `qlc-plus-swiss-knife-tool-script` project's window. A frameless
  (no-titlebar) version was tried first and reverted: testing found it
  loses native resize-by-edge and maximize on macOS, with nothing
  `pywebview` exposes to reliably bring either back.
- **"⏻ Save & Close" button** in the front end's toolbar — saves the
  currently open song, then calls the new `POST /api/quit` route to
  shut the app down cleanly. A convenience alongside the window's own
  close button, not a replacement for it.
- **`POST /api/quit`** — destroys the native window in webview mode;
  in browser mode, stops the HTTP server via `socketserver`'s own
  thread-safe `shutdown()` (not a self-sent signal — see Fixed below).
- **Clearer terminal output when `pywebview` isn't installed.** Instead
  of silently opening a browser tab, `webserver.py` now prints the
  install command (`pip install pywebview`, or `pip install -r
  requirements-optional.txt`) and reminds you of `--browser` /
  `--no-open`. Only shown when the fallback is actually due to the
  missing package — not when `--browser` was passed on purpose, and
  not with `--no-open` (nothing was going to open either way).

### Fixed
- An early draft of `/api/quit` used a self-sent `SIGINT` to stop the
  server in browser mode, matching the reference project. Caught
  during testing: self-directed signal delivery to a backgrounded
  process isn't reliable in every environment — confirmed with an
  isolated repro where even an externally sent `SIGTERM` didn't stop a
  trivial Python loop, only `SIGKILL` did. Switched to
  `socketserver.BaseServer.shutdown()`, which needs no signal delivery
  at all and works the same everywhere.

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
