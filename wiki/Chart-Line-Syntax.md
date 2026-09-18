# Chart Line Syntax

The chart line is the entire editing surface for a `render:"chart"`
section (the desktop app's editor bar, or the browser's chart-line
field for a section). It's whitespace-separated, parsed by
`grammar.py`, and every construct below round-trips: parse it, unparse
it, and you get the same line back (`grammar.unparse(grammar.parse_items(line)) == line`,
modulo whitespace).

## Symbols and frets

A bare word is a chord/note symbol:

```
B F# G E
```

Put a fret number directly in front of a symbol, no space, to record
which fret plays it:

```
5A 7D 10G
```

— fret 5 on note A, fret 7 on note D, fret 10 on note G. A leading zero
isn't allowed (`05A` is an error — write `5A`); the maximum fret is 24.

A symbol can carry a quality suffix — `m`, `maj7`, `sus4`, `/G`, and so
on are all valid: `Am`, `F#m7`, `C/G`.

## Groups and repeats

Wrap part of a line in `[ ]` and it becomes a group; put `xN` right
after the closing bracket (no space) to repeat it:

```
[5A 7D]x2
```

plays that pair twice before continuing to whatever comes next on the
line. A repeat also works standalone right after any item, block
reference, or section reference — `xN` glued on with no space.

## Annotations

A double-quoted string anywhere on the line is dropped from the
playable items and kept separately as the section's annotation (shown
in italics under the chart row):

```
G D "keep it simple" Em C
```

Since v0.20 every section card also has a visible **Annotation** field,
in every render mode. Typing `"…"` on the chart line fills that field;
clearing the field removes the annotation. Before that the annotation
could be set but never seen or removed — `unparse()` correctly leaves it
out of the line, and nothing displayed it.

## Riffs (block references)

A bare identifier that isn't a chord/note symbol is a reference to a
named riff (a `block` in the document, managed in the desktop app's
Riffs strip below the song map):

```
riff1 x3
```

Editing the riff's own content updates every section that references
it — there's only one copy of the data. A riff name can't look like a
token (e.g. `A1` is rejected as a riff name, since it would be
ambiguous with fret 1 on note A — see `model.validate_block_name`).

## Section references

`=sectionname` refers back to another section instead of repeating its
content — "same as Verse 1" rather than a byte-for-byte copy:

```
=verse1 x2
=verse1 x2 all
```

`all` (only valid after a section reference) means "repeat everything
in that section," as opposed to just its chart line. The desktop app's
Duplicate-as-reference (Ctrl/Cmd+D) inserts one of these for you.

### Names on screen, ids in the file (v0.20)

A section's **id** is minted once and never changes; its **name** can be
renamed freely. References are stored by id, so renaming a section never
breaks anything pointing at it — but an id is meaningless to read, and a
section created as "Chorus" and later renamed keeps the id `chorus1`.

So the chart line is a *view*: a reference is spelled with its target's
current name wherever it's shown, and whatever you type — `=Interlude`
or `=chorus1` — is stored as the id. Renaming a section updates every
reference to it on screen without touching a stored reference.

Matching is case-insensitive, and spaces and dashes are interchangeable
with underscores, so a section called "Verse 2" displays and resolves as
`=Verse_2`. A name that can't be written as a reference (punctuation, or
two sections sharing it) falls back to the id, which is never ambiguous.

### References are expanded when rendered (v0.20)

The pointer stays in the data — that's what makes "edit it once, every
user updates" work — but the preview and the TXT/PDF exports show the
*notes*, not `=Interlude`. A repeat expands into a bracketed group
(`[3A 12D](x2)`), a `+N`/`-N` shift expands transposed, and a reference
to a **free-text** section prints that section's text. Anything that
can't be resolved — a missing target, or a cycle — is left as a
reference rather than silently dropping the section's content.

## Transpose shift

A `+N` or `-N` glued onto a block or section reference (or a group's
closing bracket, alongside its repeat) shifts that reference's
resolved chords by N semitones without touching the referenced
riff/section itself:

```
riff1+2
=verse1-1
[5A 7D]x2+3
```

## Marks

| Typed | Means |
|---|---|
| `\|:` | repeat-open barline |
| `:\|` | repeat-close barline |
| `\|1.` | 1st ending (opens a numbered bracket over the run that follows, until the next ending mark or end of line) |
| `\|2.` | 2nd ending |
| `%` | `simile` — "play like the previous bar" |
| `rest` | rest/tacet — "don't play this measure" |
| `coda` | coda mark (renders a coda indicator; `render.has_coda()` checks for it) |
| `segno` | segno mark |
| `dc` | D.C. (da capo) |
| `ds` | D.S. (dal segno) |

## Typing it without remembering it

Each section's chart line has a palette of buttons beside it (v0.20):
`rest`, `%`, `|:`, `:|`, `|1.`, `|2.`, `x2`, `[ ]x2`, `""`, `segno`,
`coda`, `dc`, `ds`. Each drops its notation at the caret, space-separated
from what's already there, and leaves the caret inside the brackets or
quotes where that's the useful place for it. The insert fires the same
`input` event typing would, so the live parse, inline errors and preview
all behave identically — the chart line stays text you can edit by hand.

The **Notation ⌘** panel documents every mark in full: what `segno`,
`coda`, `dc` and `ds` actually tell a player to do, and the standard
D.S.–segno–coda shape they combine into.

## Sections that aren't a chart line at all

A section's render mode can be **Free** (v0.20) instead of Chart / Tab /
Both. The section becomes one plain textarea, stored verbatim as
`free_text`; nothing in it is parsed, transposed, validated or
column-aligned, and the exports print it exactly as typed. It's the
escape hatch for the bits this grammar doesn't cover.

Switching to Free is non-destructive — the section's `items` are left
alone, so switching back restores the chart line exactly. (Which is why
expanding a reference to a free section deliberately uses its *text*,
not those dormant items.)

## What a parse error looks like

A line that doesn't parse (an unmatched `[`, an invalid fret, a repeat
with nothing before it to repeat, `all` after something that isn't a
section reference, and so on) never clears what you typed. The error
shows inline, in place — the last *valid* render stays on screen above
it — in both the desktop editor bar and the browser's chart-line field.
