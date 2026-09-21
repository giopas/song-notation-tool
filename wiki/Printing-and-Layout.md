# Printing and Layout

Everything about how a chart comes out on paper, and where the files
go. Added in v0.20.

## ⚙ Layout — where the settings live

Every setting on this page is in **⚙ Layout**, at the bottom of the
sidebar (v0.24). It opens a popover rather than a modal, so the song
stays in view and the preview keeps updating while you change things,
and the line under the button says what's set ("Fit to page · Auto ·
lyrics: beside"). The desktop app has the same settings behind a
⚙ Layout button on a bar along the bottom of its window.

| Group | Setting | Stored as |
|---|---|---|
| Page | Size | `pdf_scale` |
| Page | Columns | `pdf_columns` |
| Page | Colour | `color_mode` |
| Sections | Section names | `section_layout` |
| Lyrics | Lyrics | `lyrics_layout` |
| Licks & chords | Recalled licks | `lick_refs` |
| Licks & chords | Chord shapes | `chord_sheet` |

None of them changes a single note, and they're properties of the
*song*, saved in the `.sng`, so a chart prints the same way wherever
it's opened.

## Layout — where the section name goes

| Setting | Stored as | Effect |
|---|---|---|
| Sections on top | `"banner"` (default) | A full-width heading band above each section |
| Sections on the left | `"gutter"` | The name sits in a left-hand column beside the section's first line |

The gutter saves three lines per section — for most songs the
difference between a one-page chart and a two-page one, which is what
this program exists for. Its trade-off is width: the gutter plus the
longest measure line is often what caps how far **Fit to page** can
scale, so a wide song may get bigger type with the banner layout, or in
landscape.

## Colour — a label, not a wash

| Setting | Stored as | Effect |
|---|---|---|
| Colour | `"color"` (default) | Each section *heading* is coloured by section type |
| Black & white | `"bw"` | Heading bands filled light grey, label in black |

The notes themselves are always black. Colour is there to tell you
which section you're in at a glance — tinting the notes would only make
the thing you actually read harder to read.

The palette (`constants.SECTION_COLORS`) mirrors the accent each
section card carries in the editor, darkened for white paper: a colour
that reads well on a dark background is washed out in print. Every
Verse is the same blue, every Chorus the same red, and so on by type.

Black & white uses grey fills rather than solid black on purpose. A
full-width black band per section is a lot of toner for something that
only has to say "new section starts here", and black text on grey
survives a photocopy better than white reversed out of black, which
tends to fill in.

## Size — filling the page

| Setting | Stored as | Effect |
|---|---|---|
| Fit to page | `"fit"` (default) | Scale up until the chart fills the sheet without needing another page |
| Fit to one page | `"one"` | Scale *down* as far as needed to get the whole song onto a single sheet |
| 100% / 125% / 150% / 200% | `"1.0"` … `"2.0"` | A fixed scale regardless |

**Fit to one page** inverts the constraint: for plain fit the page count
is whatever it was at 100% and the type grows into the space left over;
here the page count is fixed at one and the type is what gives. It
bisects down to a floor (`export.MIN_ONE_PAGE_SCALE`, 0.3) and stops
there — a chart that still needs two pages at the floor is printed at the
floor, because a setting called "one page" is not a licence to render a
chart too small to read from a stand.

**Fit** is computed, not guessed (`export.resolve_scale`):

- a **width bound** from the longest line actually drawn — exact, since
  the body is monospace and its geometry is linear in the type size;
- a **height bound** found by bisecting on real builds, where "fits"
  means "needs no more pages than it did at 100%".

It shrinks as well as grows: a chart whose longest line already ran off
the edge of the paper was never "fitted" by leaving it there. There's a
legibility floor (`export.MIN_FIT_SCALE`) below which it stops — past
that point the honest answer is landscape, or fewer measures per line.

The song title and the meta line are set in a proportional face, so they
aren't covered by that monospace width bound — they cap their own size
independently instead. A long title shrinks to fit rather than dragging
the whole chart's scale down with it.

Scale, colour and columns are PDF concerns; the TXT export is plain text
and ignores all three. Scale and column count are resolved *together* —
half a page's width fits a different size of type than a whole one — so
`resolve_scale()` takes the column count the document asks for.

## Columns — splitting the page

| Setting | Stored as | Effect |
|---|---|---|
| Auto | `"auto"` (default) | Two columns when they print the chart better, one when they don't |
| One column | `"1"` | The full width of the page, always |
| Two columns | `"2"` | Split, whatever it costs |

A chart is mostly short lines, so one column down an A4 page leaves half
the sheet white — and **Fit to page** can only grow the type until that
one column is full. Two columns halve the height the chart needs, and
the fit spends that height on type size instead.

The page fills top to bottom down the left column and then down the
right, with a grey rule drawn down the middle of the gutter
(`constants.PDF_COLUMN_GAP`, `PDF_COLUMN_RULE_GRAY`). The rule is not
decoration: two columns of chart with nothing between them is two charts
whose edges you have to guess, which is exactly what you can't do
mid-song. Sections are still never split — a section that won't fit in
what's left of a column starts the next one.

**Auto** (`export.resolve_columns`) doesn't guess. It resolves the scale
for one column and for two, builds both, and compares them on what
matters at a music stand, in this order:

1. **Sheets to turn.** Fewer wins. Reaching for the next page mid-song
   costs more than a slightly smaller chord symbol.
2. **Type size.** On an equal page count, two columns have to print at
   least 5% bigger (`export._AUTO_COLUMN_GAIN`) to be worth the split.

It also refuses the split outright in the two cases where half a page
simply isn't enough room: when the longest line would run over the
gutter at the chosen scale, and when the narrowest a tab block can be
drawn — the string labels plus a single measure, since a tab grid wraps
but not below one measure per line — is wider than a column.

Columns are a PDF concern. The TXT export is plain text and ignores
them, and so does the Preview pane.

## Print

**Export ▸ Print…** builds the PDF and hands it to the operating
system's default viewer, so you get the real print panel — printer,
paper size, scaling, page range — and a preview before anything reaches
paper. Deliberately not a silent `lp` job: a stage chart is exactly the
kind of thing worth eyeballing first.

In a plain browser tab there's no save panel to reach, so Print opens
the same PDF in a new tab, where Cmd/Ctrl+P does the same thing.

## Where the files are

**One file per song, named after it** (v0.26): `Artist - Title.sng`, or
just the title with no artist — the same rule in both apps
(`constants.song_file_name()`). In the browser app the file follows the
song: change the title or artist and the next Save renames it
(`SongStore.save_song()`), saying so in the toast.

- Another song is never overwritten — a taken name becomes
  `… (2).sng` — and a song is never bumped to "(2)" of itself.
- The new file is written before the old one is removed, so a failure
  part-way leaves both, never neither. Every save is atomic: written to a
  temporary file beside the song and swapped into place.
- A capitalisation-only change renames too, even on macOS's
  case-insensitive disk (where the two names are one file, so it's
  renamed in place rather than copied and deleted).
- A song with no title keeps the name it has.
- File names keep brackets, apostrophes and accented letters; anything
  that could reach outside the folder is removed.

Why not one session file holding every song? Separate files are what let
you email or AirDrop a single chart, move or back up songs one by one,
batch-export a folder from the CLI, and lose at most one song — not the
whole book — if a file is ever damaged.

**Open example song** always gives the example as shipped
(`Song Notation Tool - Example Song.sng`). An untouched one is reused,
so repeated clicks don't pile up copies; one you've edited into a song of
your own is never handed back as the example.

The desktop app's Save panel suggests the same name, but saves wherever
and as whatever you choose, and doesn't rename files on its own.

### The songs folder


Songs live in **`~/Documents/Song Notation Tool`** by default — under
your home directory, not inside the checkout, so they're covered by
whatever backs up the rest of your Documents and a `git clean` or a
fresh clone can't take them with it. `.sng` files left in an older
in-repo `songs/` folder are moved there on first run and reported by
name; nothing is ever overwritten (a name that already exists in the
target is left alone and reported).

The folder is shown in the sidebar with **Reveal** and **Change…**
buttons (native window only — a browser tab can't open a folder
picker); Change… can also move your songs, see
[Changing the songs folder](#changing-the-songs-folder) below. The
choice is remembered in:

```
~/Library/Application Support/Song Notation Tool/config.json   # macOS
%APPDATA%\Song Notation Tool\config.json                       # Windows
$XDG_CONFIG_HOME/song-notation-tool/config.json                # Linux
```

`--dir` still overrides it for a single run without changing the
stored setting.

**Exports** open a normal save panel in the native window, pre-filled
with `<Artist> - <Title>.pdf` and starting in whichever folder you
exported to last (also remembered in that config file); the app then
reports the full path it wrote. In a browser tab an export is an
ordinary download and the browser's own settings decide where it lands.

## What the printed chart shows

- **The body is 9pt Helvetica at 100%** (`export.MONO_SIZE`), headings
  10.5pt Helvetica-Bold, laid out on a **fixed column grid** 0.6 em wide
  (`MONO_CHAR_W`) — Courier's advance, which every chart row, tab grid
  and width bound is measured in. The face is proportional but the
  *positions* are monospace: see the next two points. (A Courier font
  object is declared in the PDF but not currently used.)
- **Fret numbers are a figure over the chord**, not a line of notes:
  smaller than the symbols, raised close to the line below, amber on
  colour and a light grey in black & white
  (`constants.FRET_SIZE_RATIO`, `FRET_LINE_RATIO`, `FRET_RGB`,
  `FRET_BW_RGB`). The browser's section preview sets them the same way.
- **Rows are drawn on their columns.** A chart is aligned by padding with
  spaces and counting characters, but the page is set in a proportional
  face where a space is nothing like a character wide. So each stretch of
  ink is placed at the column it was rendered at rather than the row
  being drawn as one string — the padding is measured, not drawn.
  Expanded free text is the exception: prose, not column data, so it is
  set as written.
- **Tab lines are drawn mark by mark** (v0.24.3). A lick's string line
  used to be drawn as one piece of text, and in a proportional face a
  "4" is wider than a "-": a line with more digits came out longer, and
  its frets drifted off the frets above them. Now every fret number and
  bar line sits on its own column, and each run of dashes is a solid rule
  that meets the bar lines, so a string reads as one line and frets
  played together line up across strings.
- **A named lick's name sits in front of its tab**, on its first string
  line — `Riff1 |D|-----3---|` — with the other strings indented to match
  (v0.24.2). A recalled lick printed "name only" shows as `Riff1 (x3)` on
  the chord row instead; see [[Chart Line Syntax]].
- **A tab block standing on its own** — a lick with no chord symbols
  beside it — is set off by a blank line above *and* below, so the chart
  line after it doesn't read as part of the tab. A lick written in among
  chords still sits directly under them, which is the point of writing it
  there.
- **Section names** print plain — no square brackets.
- **One instrument for the whole song** is stated once in the header,
  beside the key and tempo, rather than repeated on every section. A
  song that genuinely switches instrument keeps the per-section label.
  Filtering an export to a single instrument counts as a
  single-instrument chart and reads like one.
- **References are expanded**: a section whose content is `=Interlude`
  prints the Interlude's notes, not the pointer. See
  [[Chart Line Syntax]].
- **Footers** carry the version, the date, and the project URL.

## Lyrics — where the words go

One setting for the whole song, **⚙ Layout → Lyrics** (`lyrics_layout`,
v0.24):

| Choice | Stored as | What prints |
|---|---|---|
| Don't print | `"none"` (default for new songs) | Nothing. The words stay in the Lyrics dialog as reference |
| All together — at the start | `"start"` | Every section's words in one block before the chart, each under its section's name |
| All together — at the end | `"end"` | The same, after the chart |
| All together — left column | `"side"` | The same block down column 0 of every page, the chart in one column beside it |
| Each section — beside chart | `"beside"` | Each section's words in a column to the right of its own chart |
| Each section — under chart | `"below"` | Each section's words under its own chart |

Nothing is aligned chord to syllable in any of them. The claim is
*during these words, this is what you play*, and no more than that.

**The gathered layouts** (`start`, `end`, `side`) print
`render.lyric_blocks()`: each section's words under its name, in song
order, wrapped to the column they're given (a wrapped line hangs its
continuation further in). If no section has words yet they use the
whole-song sheet instead, with its `=== Verse 1 ===` markers promoted
to headings — the one place a marker earns ink. A per-section layout on
a song whose sheet was never split does the same at the start, so words
the song has are never silently left off the page.

**Left column** makes the page two columns whatever `pdf_columns` says,
and ⚙ Layout greys Columns out and says why. Column 0 on every page is
the words — laid out up front, page by page, and drawn as each page is
finished — so the chart and the lyrics flow independently. Words that
outrun the chart get pages of their own, which "fit to one page" counts
and shrinks the type to avoid.

**Beside the chart**: every section's words start in **one column for
the whole song** (`render.shared_lyric_column()`, v0.24.1), clear of the
widest chart that has words beside it — so down the page they read as a
column instead of zigzagging with each chart's width. A wide
instrumental section with no words doesn't push them out. The first
lyric line sits level with the chart's first line, and a section costs
the page the taller of the two sides, which is what the section-height
estimate and the width bound both measure.

**Markers never print.** `=== Verse 1 ===` lines are structure for the
editor; `render.printable_lyrics()` strips them everywhere words are
drawn.

**Older songs** keep printing what they printed: a song saved before
v0.24 opens as `"below"` if any section's old print box was ticked,
`"start"` if only the whole-song sheet's was, and `"none"` otherwise
(`model.migrate_document()`). The per-section `print_lyrics` flags no
longer decide anything.

## Chord shapes

**⚙ Layout → Chord shapes** (`chord_sheet`: `"none"` / `"start"` /
`"end"`) prints the shapes from the **Chords 🎸** dialog once, as a
block of tab-style diagrams. `chords.sheet_lines()` lays them out in as
many per row as the column width takes, grouped by instrument — a row
mixing a four-string and a six-string shape reads as one wrong diagram
rather than two right ones.

**They move with the song's transpose** (`chords.transpose_chord()`),
re-voiced the way a player would rather than slid:

| Shape as typed | Prints after a transpose as |
|---|---|
| No open strings (a barre, a power chord) | the same shape, slid — up an octave if it would fall off the nut |
| Open, and the new chord has an open shape | that open shape — C up a tone prints as `xx0232` |
| Open, no open shape for the new chord | the lower-sitting of its E-form and A-form barres |
| No barre template (add9, slash chords) | slid whole, open strings included, like a capo |

Off standard six-string tuning only the sliding rules apply. Only the
song's transpose moves the sheet, not a section's: the sheet belongs to
the whole song. The shapes stored are always the ones you typed.

## Changing the songs folder

**Change…** in the sidebar (native window, v0.25) opens the system
folder picker, then shows what would happen *before* anything does:
both paths, how many songs will move, and which will stay because a
song with the same name is already in the new folder. The choices are
**Move N songs**, **Just use this folder** (switch, leave the songs
where they are) or **Cancel**.

The move (`userpaths.relocate_songs()`) is conservative: only `.sng`
files move — the folder could be ~/Documents itself, and nothing else in
it is ours — a song already at the destination is never overwritten,
the old folder is never deleted, and the setting only changes once the
new folder is known to be usable. It works across drives and into
iCloud Drive or Dropbox. The song you have open stays open, unsaved
edits and all, if it moved; if it didn't, it's closed rather than left
for a later Save to write a stray copy into the new folder.

It's native-window only on purpose — see [[CLI and Web Use]]. In a
browser tab, `--dir` chooses the folder for a run.
