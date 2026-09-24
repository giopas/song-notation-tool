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
line. A space before the `xN` is fine too.

A repeat also works after a riff or a section reference — but there it
**needs the space**: `riff1 x3`, `=verse1 x2`. Glued on, it becomes part
of the name, so `riff1x3` looks for a riff *called* "riff1x3", and when
there isn't one it prints the bare name instead of an error. A lick
takes either form, `{Riff1}x3` or `{Riff1} x3`, since the brace ends
its name. A single chord can't carry a repeat on its own — wrap it,
`[Am]x3`.

| After | Repeat | Transpose shift |
|---|---|---|
| a group `[…]` or a lick `{…}` | `[5A 7D]x2` or `[5A 7D] x2` | not allowed |
| a riff or section reference | `riff1 x3`, `=verse1 x2` — **with** the space | `riff1 +2`, `=verse1 -1` |
| a plain chord | not allowed — `[Am]x3` | not allowed |

A group isn't limited to bare chords on one line — wrap a whole phrase,
lick or line break included, and the group repeats all of it:

```
[3G 2F# {Riff3 = G - - - - | D 4 - - - | A - 4 5 4 | E - - - -}]x4
```

A plain group of chords still prints as the single bracketed line
above — `[5A 7D]x2` — but a lick needs rows of its own for its tab, so
a group holding one is printed in full rather than collapsed down to
the lick's bare name. A right-hand bracket, spanning every row the
group prints on, carries the repeat instead of any inline `[...](xN)`
text:

```
  3  2
  G  F#                       |
         Riff3 |G|---------|  |
               |D|-4-------|  | (x4)
               |A|---4-5-4-|  |
               |E|---------|  |
```

The bracket is the point: `(x4)` sitting on its own column, next to a
bar that runs the full height of the chords *and* the tab, reads as
"this whole thing, four times" — not "the last note repeats", which is
what text tacked onto one line would say instead. The small `3  2`
above `G  F#` doesn't get its own segment of the bar — a fret number
prints as a superscript over its chord, not as a line of its own, so
the bracket starts at the chord row underneath it.

A group also isn't limited to one line for another reason: a `//`
inside it still breaks the line, exactly as it would outside any
group, and gets the same bracket treatment:

```
[7B 4G# 3G 5D // 7B 4G# 5A]x4
```

```
  7  4   3  5
  B  G#  G  D  | (x4)
  7  4   5
  B  G#  A     |
```

Each line still gets its own columns (see **Breaking a section over
several lines**, below) — the group boundary doesn't change that. The
same happens when a *reference* expands into a repeat — `=verse1 x4`
where Verse 1 itself has a `//` — since a repeated reference resolves
to this same kind of group.

## Annotations

Any of the four quote characters works — `"`, and the curly `“ ” ‘ ’`
that macOS and every word processor substitute for a typed one. A
double-quoted string anywhere on the line is dropped from the
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

Expansion is live in the editor too: changing a section — its chart
line, its free text, its render mode, or its name — refreshes every
section card that references it, so a `//` added to Chorus_1 shows up
on Chorus_2 straight away. Those cards are rendered from the stored
reference (an id) rather than from the text in the line field, so a
rename can't leave a card showing a dangling `=Old_Name`.

## Transpose shift

A `+N` or `-N`, as its own word after a riff or section reference,
shifts that reference's resolved chords by N semitones without touching
the referenced riff/section itself. With a repeat, the shift comes
after it:

```
riff1 +2
=verse1 -1
riff1 x2 +3
```

A shift isn't valid after a group or a lick — to move a group, put it in
a riff and shift the reference, or transpose the whole section.

## Breaking a section over several lines

A long section reads better grouped into the phrases you'd write out by
hand. `//` anywhere on the chart line starts a new line at that point:

```
5A 5D // 8F 5D 7E // 5A 8F 5D 7E // 5A 5A 8F
```

renders as

```
  5  5
  A  D
  8  5  7
  F  D  E
  5  8  5  7
  A  F  D  E
  5  5  8
  A  A  F
```

Each block aligns its own columns — that's the point of breaking, rather
than having one grid stretch across the whole section. In the desktop
song map, where the section name sits in a left gutter, the continuation
lines stack under the first one rather than sliding back to the margin.

Adding `>` pushes that line in — `//>` for one step, `//>>` for two — so
a phrase can sit visibly inside the one above it:

```
5A 5D //> 8F 5D 7E // 5A 5A 8F
```

```
  5  5
  A  D
        8  5  7
        F  D  E
  5  5  8
  A  A  F
```

The indent is relative to wherever the line would otherwise start, so it
reads the same in the export and in the desktop song map, where lines
already sit in the section-name gutter. One step is
`render.INDENT_STEP` characters.

It's purely a layout mark: it plays nothing, it's stored as a
`line_break` mark like any other, and the chart line itself stays a
single line you can edit. `//` rather than a bare `/` because a slash
already lives inside chord symbols (`C/G`) — as a whole word a single
slash would be unambiguous, but doubling it keeps it obvious which is
meant.

## Licks — a bit of tab, in among the chords

Some things aren't a chord. When you want to remember the actual notes of
a short figure, write it in braces at the point it's played:

```
C {G 5 7 5 | D - - 3} G Am
```

Each `|`-separated line is one string — its name, then its frets.
Positions align by column across the lines, so a fret on `G` at the third
position sounds with whatever is third on `D`. `-` means that string
isn't played there and `x` means muted. Write as many lines as the figure
needs.

A `G: 5 7 5` spelling is accepted too, if the colon reads better.

It renders as a small tab block sitting under its own column in the chart
row:

```
  C           G  Am
     G 5 7 5
     D - - 3
```

— so it stays where it belongs in the sequence, rather than being exiled
to a separate tab grid.

It **prints as tab**, one row per string, with the string named in its own
little box at the left:

```
|G|-5-7-5-|
|D|-----3-|
```

Every position is the same width down the whole lick, so the columns line
up vertically and you read it the way you read a tab staff. On screen and
in a colour PDF the lick is **blue** — it isn't a chord symbol and
shouldn't be read as one at a glance from a stand.

The `{ tab }` button beside the chart line inserts an empty one: every
string of that section's instrument, six positions wide, all dashes. Type
frets over the dashes; delete nothing. Widen or narrow it by adding or
removing dashes — a lick is as long as you write it.

### Naming a lick

Put a name and an `=` at the front and the lick has a handle:

```
{Riff1 = G 5 7 5 | D - - 3}
```

It prints with the name in front of the tab, on its first string line —
`Riff1 |G|-5-7-5-|` — and anywhere else in the song `{Riff1}` plays it
again, with a repeat if you want one:

```
Intro    {Riff1}
Chorus   {Riff1}x3   5A 5D
```

A reference resolves at render time to the notes themselves, with the
name (and repeat) in front of them — `Riff1 (x3) |G|…`. So editing the
lick updates every place that plays it, exactly as a section reference
does — and there is no separate library to keep in step: the name lives
on the lick where you wrote it.

A lick doesn't have to stand alone to get a repeat — wrap it together
with whatever comes before it and repeat the group instead (see
**Groups and repeats**, above) when it's the *pairing* that repeats,
not just the riff:

```
[3B 2F# {Riff3 = G - - - - | D 4 - - - | A - 4 5 4 | E - - - -}]x4
```

**Braces recall a lick; `=` recalls a section.** `{Riff1}` plays the
lick named Riff1. `=Riff1` looks for a *section* called Riff1, and when
there isn't one it prints as the literal text `=Riff1`. Braces define a
lick and braces recall it.

If you'd rather the notes printed once — where the lick was written —
and every recall said just `Riff1 (x3)`, set **⚙ Layout → Recalled
licks → Name only** (`lick_refs: "name"`; the default is `"tab"`). A
reference that doesn't find its lick keeps its braces on paper, so a
misspelt name is visible. Either way a lick's name prints in the lick
blue.

Names match loosely (`{riff1}` finds `Riff1`) and can't start like a
chord: `A1` would read as a chord symbol on a chart line, so it's
refused. The `{ name = tab }` button fills a free name in for you —
`Riff1`, `Riff2`, and so on — and in the browser every lick you've
already named appears in the palette as a button that inserts the
reference. (A lick named while you type gets its button once the
section card is next redrawn — after Save, for instance.)

Transposing the section moves a lick's frets with
it: same strings, shifted positions. A fret that would fall off either
end of the neck is left as it was rather than silently clamped to a
position you'd actually play.

## Marks

| Typed | Means |
|---|---|
| `\|:` | repeat-open barline |
| `:\|` | repeat-close barline |
| `\|1.` | 1st ending (opens a numbered bracket over the run that follows, until the next ending mark or end of line) |
| `\|2.` | 2nd ending |
| `//` | line break — start a new line here (layout only; plays nothing) |
| `//>` | line break, indented one step per `>` |
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
