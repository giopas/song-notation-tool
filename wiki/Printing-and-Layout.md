# Printing and Layout

Everything about how a chart comes out on paper, and where the files
go. Added in v0.20.

The three settings below live in the **Print & export** group on the
second row of the meta panel, framed separately from Title/Artist/Key
because none of them changes a single note. They're properties of the
*song*, saved in the `.sng` (`section_layout`, `color_mode`,
`pdf_scale`), so a chart prints the same way wherever it's opened.

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
| 100% / 125% / 150% / 200% | `"1.0"` … `"2.0"` | A fixed scale regardless |

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

Scale and colour are PDF concerns; the TXT export is plain text and
ignores both.

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
