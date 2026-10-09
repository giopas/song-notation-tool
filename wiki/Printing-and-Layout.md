# Printing and layout

How a chart comes out on paper, and where the files go.

## ⚙ Layout: where the settings are

Every setting on this page is in **⚙ Layout**, at the bottom of the
sidebar. It opens a popover instead of a modal window, so the song stays
in view and the preview keeps updating while you change things. The line
under the button shows what is set ("Fit to page · Auto · lyrics:
beside"). The desktop app has the same settings behind a **⚙ Layout**
button on a bar along the bottom of its window.

| Group | Setting | Stored as |
|---|---|---|
| Page | Size | `pdf_scale` |
| Page | Columns | `pdf_columns` |
| Page | Colour | `color_mode` |
| Sections | Section names | `section_layout` |
| Lyrics | Lyrics | `lyrics_layout` |
| Licks & chords | Recalled licks | `lick_refs` |
| Licks & chords | Chord shapes | `chord_sheet` |
| Instruments | Which instruments print | `print_instruments` |

None of them changes a note. They belong to the song and are saved in the
`.sng`, so a chart prints the same wherever it is opened.

## Parts and instruments

A section can hold more than one instrument. The section's own chart line
belongs to the instrument picked in its header; click **+ Part** under it
to add another instrument's line to the same section, for example the
guitar chords under the bass line of a verse. A part has its own
instrument and its own chart line, with the same grammar, and plays the
section's repeats and transpose.

A part prints under the section's own line, with its instrument as a small
grey heading. **⚙ Layout → Instruments** picks which instruments print: all
of them (nothing ticked), or only the ones you tick. A section prints when
its own instrument or one of its parts is picked, so a guitarist's copy
can leave out the bass lines and keep the guitar parts. A choice that
matches nothing in the song is ignored, so the chart never comes out
empty because of it. The command line's `--instrument` does the same for
one export.

Parts are edited in the browser app. The desktop app lists them under
their section and keeps them as they are when it saves.

## Where the section name goes

| Setting | Stored as | Effect |
|---|---|---|
| Sections on top | `"banner"` (default) | A full-width heading band above each section |
| Sections on the left | `"gutter"` | The name sits in a left-hand column beside the section's first line |

The left column saves three lines per section. For most songs that is the
difference between a one-page and a two-page chart, and a one-page chart
is what this program is for. The cost is width: the column plus the
longest measure line is often what limits how far **Fit to page** can
scale, so a wide song may get bigger type with names on top, or in
landscape.

## Colour labels the sections

| Setting | Stored as | Effect |
|---|---|---|
| Colour | `"color"` (default) | Each section heading is coloured by section type |
| Black & white | `"bw"` | Heading bands filled light grey, label in black |

The notes are always black. The colour tells you at a glance which section
you are in; colouring the notes would make the part you read harder to
read.

The palette (`constants.SECTION_COLORS`) follows the colour each section
card has in the editor, darkened for white paper, because a colour that
reads well on a dark background looks washed out in print. Every Verse is
the same blue, every Chorus the same red, and so on by type.

Black & white uses grey fills on purpose, not solid black. A full-width
black band for every section uses a lot of toner just to say "new section
starts here", and black text on grey survives a photocopy better than
white text on black, which tends to fill in.

## Size: filling the page

| Setting | Stored as | Effect |
|---|---|---|
| Fit to page | `"fit"` (default) | Scale up until the chart fills the sheet without needing another page |
| Fit to one page | `"one"` | Scale down as far as needed to get the whole song onto one sheet |
| 100% / 125% / 150% / 200% | `"1.0"` … `"2.0"` | A fixed scale |

Fit to one page turns the rule around. With plain Fit to page, the page
count is whatever it was at 100% and the type grows into the space left
over. With Fit to one page, the page count is fixed at one and the type
size gives way. It searches down to a floor (`export.MIN_ONE_PAGE_SCALE`,
0.3) and stops there: a chart that still needs two pages at the floor is
printed at the floor, so the setting never makes a chart too small to read
on a stand.

Fit to page is calculated (`export.resolve_scale`):

- a width limit from the longest line drawn. This is exact, because the
  body is monospaced and its size grows in proportion to the type size;
- a height limit found by bisecting on real builds, where "fits" means
  "needs no more pages than at 100%".

It shrinks as well as grows, so a chart whose longest line ran off the
edge of the paper gets fitted too. Below a readability floor
(`export.MIN_FIT_SCALE`) it stops; past that point the answer is
landscape, or fewer measures per line.

The song title and the details line are set in a proportional font, so
the monospaced width limit does not cover them. They limit their own size
instead, and a long title shrinks to fit without pulling the whole chart's
scale down.

Scale, colour and columns only apply to the PDF; the TXT export is plain
text and ignores all three. Scale and the number of columns are worked out
together (half a page's width fits a different type size than a whole
one), so `resolve_scale()` takes the column count the song asks for.

## Columns: splitting the page

| Setting | Stored as | Effect |
|---|---|---|
| Auto | `"auto"` (default) | Two columns when they print the chart better, one when they do not |
| One column | `"1"` | The full width of the page, always |
| Two columns | `"2"` | Always split |

A chart is mostly short lines, so one column down an A4 page leaves half
the sheet white, and Fit to page can only grow the type until that one
column is full. Two columns halve the height the chart needs, and the fit
spends that height on type size.

The page fills from top to bottom down the left column, then down the
right, with a grey rule down the middle of the gutter
(`constants.PDF_COLUMN_GAP`, `PDF_COLUMN_RULE_GRAY`). The rule matters: two
columns of chart with nothing between them look like two charts whose
edges you have to guess, and you cannot afford that mid-song. Sections are
still never split; a section that does not fit in what is left of a column
starts the next one.

Auto (`export.resolve_columns`) works out the scale for one column and for
two, builds both, and compares them on what matters at a music stand, in
this order:

1. Sheets to turn. Fewer wins: reaching for the next page mid-song costs
   more than a slightly smaller chord symbol.
2. Type size. With the same number of pages, two columns have to print at
   least 5% bigger (`export._AUTO_COLUMN_GAIN`) to be worth the split.

It never splits in the two cases where half a page is too narrow: when the
longest line would run over the gutter at the chosen scale, and when the
narrowest possible tab block (the string labels plus a single measure; a
tab grid wraps, but not below one measure per line) is wider than a
column.

The TXT export is plain text and ignores columns, and so does the Preview.

## Print

**Export** → **Print…** builds the PDF and opens it in the system's default
viewer, so you get the real print panel (printer, paper size, scaling, page
range) and a preview before anything is printed. It is deliberately not a
silent `lp` job: a stage chart is worth checking before it goes on paper.

A plain browser tab has no save panel to reach, so **Print…** opens the
same PDF in a new tab, where Cmd/Ctrl+P does the same.

## Where the files are

Each song is its own file, named after the song: `Artist - Title.sng`, or
just the title if there is no artist. Both apps use the same rule
(`constants.song_file_name()`). In the browser app the file follows the
song: change the title or artist and the next Save renames it
(`SongStore.save_song()`), and the message after saving says so.

- Another song is never overwritten (a name that is taken becomes
  `… (2).sng`), and a song is never renamed to "(2)" of itself.
- The new file is written before the old one is removed, so a failure
  halfway leaves both files, never neither. Every save is atomic: written
  to a temporary file next to the song, then moved into place.
- A change of capital letters only renames the file too, even on macOS's
  case-insensitive disks, where the two names are one file and it is
  renamed in place instead of copied and deleted.
- A song with no title keeps the file name it has.
- File names keep brackets, apostrophes and accented letters. Anything
  that could point outside the folder is removed.

Separate files, instead of one session file for every song, let you email
or AirDrop a single chart, move or back up songs one at a time,
batch-export a folder from the CLI, and lose at most one song, not the
whole book, if a file is ever damaged.

**Open example song** always gives the example as shipped
(`Song Notation Tool - Example Song.sng`). An untouched copy is reused, so
repeated clicks do not pile up copies, and one you have edited into a song
of your own is never handed back as the example.

The desktop app's save panel suggests the same name, but saves wherever
and under whatever name you choose, and does not rename files by itself.

### The songs folder

Songs live in `~/Documents/Song Notation Tool` by default, under your home
folder and outside the checkout. Whatever backs up your Documents covers
them, and a `git clean` or a fresh clone cannot remove them. `.sng` files
left in an older in-repo `songs/` folder are moved there on first run and
listed by name. Nothing is overwritten: a name that already exists in the
target folder is left alone and reported.

The sidebar shows the folder with **Reveal** and **Change…** buttons. They
are in the native window only, because a browser tab cannot open a folder
picker. **Change…** can also move your songs; see
[Changing the songs folder](#changing-the-songs-folder). The choice is saved
in:

```
~/Library/Application Support/Song Notation Tool/config.json   # macOS
%APPDATA%\Song Notation Tool\config.json                       # Windows
$XDG_CONFIG_HOME/song-notation-tool/config.json                # Linux
```

`--dir` overrides it for one run without changing the saved setting.

Exports open the system's save panel in the native window, filled in with
`<Artist> - <Title>.pdf` and starting in the folder you last exported to
(also saved in that config file). The app then shows the full path it
wrote. In a browser tab an export is an ordinary download, and the
browser's settings decide where it goes.

### How an export is written

Every export, from any front end, goes through `export.py`:

- The song is checked first (`check_document()`, the same check as
  `cli.py lint`). A chart line that cannot be written out or does not
  parse again stops the export, and the message names the section.
  Chord-shape notes (chords with no shape, shapes never played) are shown
  but never stop it.
- The bytes come from one place (`export_bytes()`): TXT is UTF-8 with Unix
  line endings and a newline at the end on every system, and PDF is
  `build_pdf()`. A download, a saved file and a command-line export of the
  same song are the same bytes.
- The file is written to a temporary file in the same folder and then
  moved into place, so a crash halfway never leaves half a file.
- An existing file is replaced only when you agreed to it in a save
  dialog. Otherwise (the command line without `--force`) the export goes
  to the next free name, `song_v2.pdf`.

## What the printed chart shows

- The body is 9 pt Helvetica at 100% (`export.MONO_SIZE`), and headings are
  10.5 pt Helvetica-Bold. Everything is placed on a fixed column grid 0.6
  em wide (`MONO_CHAR_W`, Courier's character width), which every chart
  row, tab grid and width limit is measured in. The font is proportional
  but the positions are monospaced, as the next two points explain. (The
  PDF declares a Courier font object, which is not used at the moment.)
- Fret numbers are a small figure over the chord, not a line of notes:
  smaller than the symbols, raised close to the line below, amber in colour
  and light grey in black and white (`constants.FRET_SIZE_RATIO`,
  `FRET_LINE_RATIO`, `FRET_RGB`, `FRET_BW_RGB`). The browser's section
  preview sets them the same way.
- Rows are drawn on their columns. A chart is lined up by padding with
  spaces and counting characters, but the page uses a proportional font
  where a space is much narrower than a character. So each piece of text
  is placed at the column it was rendered at, instead of the row being
  drawn as one string: the padding is measured, not drawn. Expanded free
  text is the exception, because it is prose and not column data, so it is
  set as written.
- Tab lines are drawn mark by mark. Every fret number and bar line sits on
  its own column, and each run of dashes is a solid line that meets the
  bar lines, so a string reads as one line and frets played together line
  up across strings, even though a "4" is wider than a "-" in a
  proportional font.
- A named lick's name sits in front of its tab, on its first string line
  (`Riff1 |D|-----3---|`), with the other strings indented to match. A
  recalled lick printed "name only" shows as `Riff1 (x3)` on the chord row
  instead; see [[Chart Line Syntax]].
- A tab block on its own (a lick with no chord symbols beside it) gets a
  blank line above and below, so the chart line after it does not read as
  part of the tab. A lick written among chords sits directly under them,
  which is why you write it there.
- Section names print plain, without square brackets.
- When the whole song uses one instrument, it is printed once in the
  header, next to the key and tempo, instead of on every section. A song
  that really switches instrument keeps the label on each section. An
  export filtered to one instrument counts as a single-instrument chart
  and prints like one.
- References are expanded: a section whose content is `=Interlude` prints
  the Interlude's notes, not the pointer. See [[Chart Line Syntax]].
- Endings print as brackets: a line over the bars with a short drop at
  the start and the number ("1.", "2.") in bold. Palm mute prints "P.M."
  and a dashed line ending in a tick, cresc. and dim. print as hairpins,
  and an accent as a small wedge over its note. These sit on a row of
  their own above the fret numbers.
- A coda prints as a box with the coda sign and "Coda" in bold, so the
  place to jump to is easy to find on a stand. Dynamics print in bold.
- The footer shows the version, the date and the project URL.
- Every PDF carries the song it was made from, attached as `song.sng`.
  Open the PDF with **Import song…** in the browser app, **Open** in the
  desktop app, or `cli.py extract`, and you get the song back to edit. A
  PDF viewer lists it as an attachment. The TXT export has nowhere to keep
  it, so only a PDF can be reopened. PDFs made before 0.30 have no song
  attached.

## Where the words go

One setting for the whole song, **⚙ Layout** → Lyrics (`lyrics_layout`):

| Choice | Stored as | What prints |
|---|---|---|
| Don't print | `"none"` (default for new songs) | Nothing. The words stay in the Lyrics dialog for reference |
| All together — at the start | `"start"` | Every section's words in one block before the chart, each under its section's name |
| All together — at the end | `"end"` | The same, after the chart |
| All together — left column | `"side"` | The same block down column 0 of every page, with the chart in one column beside it |
| Each section — beside chart | `"beside"` | Each section's words in a column to the right of its own chart |
| Each section — under chart | `"below"` | Each section's words under its own chart |

None of them lines up chords with syllables. What the chart says is
"during these words, this is what you play", and no more.

The "all together" layouts (`start`, `end`, `side`) print
`render.lyric_blocks()`: each section's words under its name, in song
order, wrapped to the column they get (a wrapped line indents its
continuation). If no section has words yet, they use the whole-song sheet
instead, with its `=== Verse 1 ===` markers turned into headings, the only
place a marker is printed. A per-section layout on a song whose sheet was
never split does the same at the start, so words the song has never go
missing from the page.

Left column makes the page two columns whatever `pdf_columns` says, and
**⚙ Layout** greys out Columns and says why. Column 0 on every page holds
the words, laid out first, page by page, and drawn as each page is
finished, so the chart and the lyrics flow separately. Words that run
longer than the chart get pages of their own, which Fit to one page counts
and shrinks the type to avoid.

Beside the chart, every section's words start in one column for the whole
song (`render.shared_lyric_column()`), clear of the widest chart that has
words beside it, so they read down the page as a column instead of
zigzagging with each chart's width. A wide instrumental section with no
words does not push them out. The first line of words is level with the
chart's first line, and a section takes the height of the taller of the
two sides, which is what the section-height estimate and the width limit
both measure.

Markers never print. `=== Verse 1 ===` lines are structure for the editor,
and `render.printable_lyrics()` removes them wherever words are drawn.

Older songs keep printing what they printed: a song saved before v0.24
opens as `"below"` if any section's old print box was ticked, `"start"` if
only the whole-song sheet's box was, and `"none"` otherwise
(`model.migrate_document()`). The per-section `print_lyrics` flags no
longer decide anything.

## Chord shapes

**⚙ Layout** → Chord shapes (`chord_sheet`: `"none"`, `"start"` or `"end"`)
prints the shapes from the **Chords 🎸** dialog once, as a block of
tab-style diagrams. `chords.sheet_lines()` puts as many in a row as the
column width allows, grouped by instrument, because a row mixing a
four-string and a six-string shape reads as one wrong diagram instead of
two right ones.

The shapes follow the song's transpose (`chords.transpose_chord()`) and
are re-voiced as a player would, not just slid:

| Shape as typed | Prints after a transpose as |
|---|---|
| No open strings (a barre, a power chord) | the same shape, slid, and up an octave if it would fall off the nut |
| Open, and the new chord has an open shape | that open shape: C up a tone prints as `xx0232` |
| Open, and the new chord has no open shape | the lower of its E-form and A-form barres |
| No barre template (add9, slash chords) | slid whole, open strings included, like a capo |

On anything other than standard six-string tuning only the sliding rules
apply. Only the song's transpose moves the shapes, not a section's,
because the shapes belong to the whole song. The stored shapes are always
the ones you typed.

## Changing the songs folder

**Change…** in the sidebar (native window) opens the system folder picker,
then shows what would happen before anything does: both paths, how many
songs will move, and which will stay because a song with the same name is
already in the new folder. The choices are **Move N songs**, **Just use
this folder** (switch, and leave the songs where they are) or **Cancel**.

The move (`userpaths.relocate_songs()`) is careful. Only `.sng` files move,
because the folder could be ~/Documents itself and nothing else in it
belongs to the app. A song already at the destination is never
overwritten, the old folder is never deleted, and the setting only changes
once the new folder is known to work. It works across drives and into
iCloud Drive or Dropbox. If the song you have open moved, it stays open,
unsaved edits included. If it did not move, it is closed, so a later Save
cannot write a stray copy into the new folder.

It is in the native window only on purpose; see [[CLI and Web Use]]. In a
browser tab, `--dir` chooses the folder for a run.
