# Printing and Layout

Everything about how a chart comes out on paper, and where the files
go. Added in v0.20.

The four settings below live in the **Print & export** group on the
second row of the meta panel, framed separately from Title/Artist/Key
because none of them changes a single note. They're properties of the
*song*, saved in the `.sng` (`section_layout`, `color_mode`,
`pdf_scale`, `pdf_columns`), so a chart prints the same way wherever
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

Songs live in **`~/Documents/Song Notation Tool`** by default — under
your home directory, not inside the checkout, so they're covered by
whatever backs up the rest of your Documents and a `git clean` or a
fresh clone can't take them with it. `.sng` files left in an older
in-repo `songs/` folder are moved there on first run and reported by
name; nothing is ever overwritten (a name that already exists in the
target is left alone and reported).

The folder is shown in the sidebar with **Reveal** and **Change…**
buttons (native window only — a browser tab can't open a folder
picker). The choice is remembered in:

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

- **The body is 9pt Courier at 100%** (`export.MONO_SIZE`), headings
  10.5pt Helvetica-Bold. Size and character width are locked together:
  Courier's advance is exactly 0.6 em and that advance is the column
  arithmetic every chart row, tab grid and width bound is measured in, so
  `MONO_CHAR_W` follows `MONO_SIZE` and neither moves alone.
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
- **Lyrics print beside their chart**, not under it (`lyrics_layout`,
  `"beside"` by default; `"below"` for the old behaviour). The chart keeps
  a narrow left column, the words run down their own to the right of it,
  starting level with the chart's first line — `render.lyric_column_x()`
  picks the column, `render.compose_beside()` lays the two out for TXT,
  and the PDF draws them as two independent runs from the same starting
  `y` so a rest or a fret row keeps its own line height. Nothing is
  aligned chord to syllable: the claim is *during these words, this is
  what you play*, and no more than that. A verse therefore costs the page
  the taller of the two sides — which is what the section-height estimate
  and the width bound both measure.
- **Lyric section markers never print.** `=== Verse 1 ===` lines are
  structure for the editor; `render.printable_lyrics()` strips them
  everywhere words are drawn.
- **Chord shapes print once**, as a block of tab-style diagrams at the
  start or the end of the chart (`chord_sheet`: `"none"` / `"start"` /
  `"end"`). `chords.sheet_lines()` lays them out in as many per row as
  the column width takes, grouped by instrument — a row mixing a
  four-string and a six-string shape reads as one wrong diagram rather
  than two right ones.
- **Chord shapes transpose with the song** (`chords.transpose_chord()`),
  re-voiced rather than slid: a movable shape slides; an open shape goes
  to the new chord's open shape if standard tuning has one, else to the
  lower of its E-form and A-form barres; anything with no template is
  slid whole, like a capo. Only the song's transpose applies — the sheet
  belongs to the whole song, not to any one section.
- **Lyric placement is one song-wide setting** (`lyrics_layout`, v0.24):
  `none`, `start`, `end`, `side`, `beside`, `below`. The gathered modes
  (`start`/`end`/`side`) print `render.lyric_blocks()` — each section's
  words under its name, or the marked-up sheet if no section has any —
  wrapped to the column they're given. `side` makes the page two columns
  whatever `pdf_columns` says: column 0 on every page is the words, laid
  out up front page by page and drawn as each page is finished, so the
  chart and the lyrics flow independently; words that outrun the chart
  add pages, which the one-page fit counts and shrinks to avoid.
- **Every print setting lives in ⚙ Layout**, bottom left, in both front
  ends.
