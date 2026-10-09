# Chart line syntax

For a `render:"chart"` section, the chart line is all there is to edit:
the desktop app's editor bar, or a section's chart-line field in the
browser. Words are separated by spaces and parsed by `grammar.py`. Every
form below comes back unchanged when you parse it and write it out again
(`grammar.unparse(grammar.parse_items(line)) == line`, apart from
whitespace).

## Symbols and frets

A plain word is a chord or note symbol:

```
B F# G E
```

Put a fret number right in front of a symbol, with no space, to record
which fret plays it:

```
5A 7D 10G
```

That is fret 5 on A, fret 7 on D and fret 10 on G. A leading zero is not
allowed (`05A` is an error; write `5A`), and the highest fret is 24.

A symbol can have a quality suffix such as `m`, `maj7`, `sus4` or `/G`:
`Am`, `F#m7`, `C/G`.

## Groups and repeats

Put part of a line in `[ ]` to make a group, and put `xN` right after the
closing bracket (no space) to repeat it:

```
[5A 7D]x2
```

plays the pair twice, then carries on with the rest of the line. A space
before the `xN` works too.

A repeat also works after a riff or a section reference, but there it
needs the space: `riff1 x3`, `=verse1 x2`. Without the space it becomes
part of the name, so `riff1x3` looks for a riff called "riff1x3", and if
there is none it prints the name as it is instead of showing an error. A
lick takes either form, `{Riff1}x3` or `{Riff1} x3`, because the brace
ends its name. A single chord cannot carry a repeat by itself; put it in
a group, `[Am]x3`.

| After | Repeat | Transpose shift |
|---|---|---|
| a group `[…]` or a lick `{…}` | `[5A 7D]x2` or `[5A 7D] x2` | not allowed |
| a riff or section reference | `riff1 x3`, `=verse1 x2`, with the space | `riff1 +2`, `=verse1 -1` |
| a plain chord | not allowed; use `[Am]x3` | not allowed |

A group can hold more than chords on one line. Wrap a whole phrase,
including a lick or a line break, and the group repeats all of it:

```
[3G 2F# {Riff3 = G - - - - | D 4 - - - | A - 4 5 4 | E - - - -}]x4
```

A group of plain chords still prints on one bracketed line, `[5A 7D]x2`.
A lick needs rows of its own for its tab, so a group that holds one is
printed in full, and a bracket on the right, running down every row the
group prints on, carries the repeat:

```
  3  2
  G  F#                       │
         Riff3 |G|---------|  │
               |D|-4-------|  │ (x4)
               |A|---4-5-4-|  │
               |E|---------|  │
```

With `(x4)` centred against a bar that covers the chords and the tab, the
whole phrase reads as played four times. Text at the end of one line would
look as if only the last note repeats. The small `3  2` above `G  F#` gets
no part of the bar: a fret number prints as a superscript over its chord,
not as a line of its own, so the bracket starts at the chord row below it.

The PDF and the web app's live preview draw the bracket as one solid line,
so it stays continuous across the gap of a fret row. TXT export and the
desktop app have no vector drawing, so there it is the character version
above: a `|` down the right of each row it covers, with `(xN)` on the row
nearest the middle. A fret row never gets a `|` in either version.

A `//` inside a group also breaks the line, as it does outside a group,
and gets the same bracket:

```
[7B 4G# 3G 5D // 7B 4G# 5A]x4
```

In the PDF and the web preview:

```
  7  4   3  5
  B  G#  G  D  │
  7  4   5     │ (x4)
  B  G#  A     │
```

In TXT export and the desktop app:

```
  7  4   3  5
  B  G#  G  D  | (x4)
  7  4   5
  B  G#  A     |
```

Each line still lines up its own columns (see
[Breaking a section over several lines](#breaking-a-section-over-several-lines)),
whether or not it is in a group. The same happens when a reference
expands into a repeat, such as `=verse1 x4` where Verse 1 has a `//`,
because a repeated reference becomes this same kind of group.

## Annotations

Any of the four quote characters works: `"`, and the curly `“ ” ‘ ’` that
macOS and word processors put in place of a typed one. A quoted string
anywhere on the line is taken out of the playable items and kept as the
section's annotation, shown in italics under the chart row:

```
G D "keep it simple" Em C
```

Every section card also has an Annotation field, in every render mode.
Typing `"…"` on the chart line fills that field, and clearing the field
removes the annotation.

## Riffs (block references)

A plain word that is not a chord or note symbol refers to a named riff (a
`block` in the document). Riffs are managed in the **Riffs 🔁** dialog in
the browser app, or in the Riffs strip below the song map in the desktop
app:

```
riff1 x3
```

Editing the riff changes every section that uses it, because there is
only one copy. A riff name cannot look like a chord token: `A1` is refused
as a riff name, because it would read as fret 1 on A (see
`model.validate_block_name`). Dynamics (`mf`) and the run words (`pm`,
`cresc`, `dim`) are refused for the same reason.

In the **Riffs 🔁** dialog you can create a riff from a chart line, change
its line, rename it (every line that plays it follows the new name) and
delete it once nothing plays it. To make a riff from something already
typed, select that part of a section's chart line and click **→ riff**
beside it: the selection becomes `riff1` and the riff holds what you
selected.

## Section references

`=sectionname` points back to another section instead of repeating its
content, the way you would write "same as Verse 1":

```
=verse1 x2
=verse1 x2 all
```

`all` (only after a section reference) repeats everything in that section,
not only its chart line. **Duplicate as reference** (Ctrl/Cmd+D) in the
desktop app inserts one for you.

### Names on screen, ids in the file

A section gets an id once, and the id never changes. Its name can change
as often as you like. References are stored by id, so renaming a section
never breaks anything that points to it. An id is not meant to be read,
though: a section created as "Chorus" and renamed later keeps the id
`chorus1`.

So the chart line shows a reference by its target's current name, and
whatever you type, `=Interlude` or `=chorus1`, is stored as the id.
Renaming a section updates every reference to it on screen without
changing what is stored.

Matching ignores case, and spaces and dashes count as underscores, so a
section called "Verse 2" shows and resolves as `=Verse_2`. A name that
cannot be written as a reference (punctuation, or two sections with the
same name) falls back to the id, which is never ambiguous.

### References are expanded when rendered

The data keeps the pointer, which is how "edit it once, every use
changes" works, but the preview and the TXT and PDF exports show the
notes instead of `=Interlude`. A repeat expands into a bracketed group
(`[3A 12D](x2)`), a `+N` or `-N` shift expands transposed, and a reference
to a free-text section prints that section's text. A reference that
cannot be resolved (a missing target, or a loop) is left as a reference,
so the section's content is never dropped without notice.

The editor expands references live too. Changing a section's chart line,
free text, render mode or name redraws every card that refers to it, so a
`//` added to Chorus_1 shows on Chorus_2 at once. Those cards are drawn
from the stored reference (an id) and not from the text in the line field,
so a rename cannot leave a card showing an outdated `=Old_Name`.

## Transpose shift

A `+N` or `-N`, as its own word after a riff or section reference, shifts
that reference's chords by N semitones without changing the riff or
section itself. With a repeat, the shift comes after it:

```
riff1 +2
=verse1 -1
riff1 x2 +3
```

A shift is not allowed after a group or a lick. To move a group, put it in
a riff and shift the reference, or transpose the whole section.

## Breaking a section over several lines

A long section reads better split into the phrases you would write by
hand. `//` anywhere on the chart line starts a new line there:

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

Each line lines up its own columns, so one grid does not stretch across
the whole section. In the desktop song map, where the section name sits in
a left gutter, the extra lines stack under the first one instead of going
back to the margin.

Adding `>` indents the line: `//>` for one step, `//>>` for two, so a
phrase can sit inside the one above it:

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

The indent counts from wherever the line would otherwise start, so it
looks the same in the export and in the desktop song map, where lines
already sit in the section-name gutter. One step is `render.INDENT_STEP`
characters.

The break only affects layout. It plays nothing, it is stored as a
`line_break` mark like any other, and the chart line itself stays one line
you can edit. It is `//` and not `/` because a slash already appears inside
chord symbols (`C/G`). A single slash as a word of its own would still be
clear, but the double one makes it obvious which is meant.

## Licks: a bit of tab among the chords

Some things are not chords. To remember the actual notes of a short
figure, write it in braces where it is played:

```
C {G 5 7 5 | D - - 3} G Am
```

Each part between `|` is one string: its name, then its frets. Positions
line up by column across the strings, so a fret third on `G` sounds with
whatever is third on `D`. `-` means the string is not played there, and
`x` means muted. Write as many strings as the figure needs.

`G: 5 7 5`, with a colon, is accepted too, if you find it easier to read.

On screen it is a small tab block under its own column in the chart row:

```
  C           G  Am
     G 5 7 5
     D - - 3
```

so it stays in its place in the sequence, and does not have to go in a
separate tab grid.

It prints as tab, one row per string, with the string name in a small box
on the left:

```
|G|-5-7-5-|
|D|-----3-|
```

Every position has the same width down the whole lick, so the columns
line up and you read it like a tab staff. On screen and in a colour PDF
the lick is blue, so it is not mistaken for a chord symbol when you glance
at it on a stand.

The **{ tab }** button beside the chart line inserts an empty lick: every
string of the section's instrument, six positions wide, all dashes. Type
frets over the dashes without deleting anything. Add or remove dashes to
make it longer or shorter; a lick is as long as you write it.

### Naming a lick

Put a name and `=` at the front, and the lick has a name you can call:

```
{Riff1 = G 5 7 5 | D - - 3}
```

It prints with the name in front of the tab, on its first string line
(`Riff1 |G|-5-7-5-|`), and `{Riff1}` anywhere else in the song plays it
again, with a repeat if you like:

```
Intro    {Riff1}
Chorus   {Riff1}x3   5A 5D
```

When the chart is rendered, a reference becomes the notes themselves, with
the name (and the repeat) in front: `Riff1 (x3) |G|…`. Editing the lick
changes every place that plays it, as with a section reference, and there
is no separate library to keep up to date: the name lives on the lick
where you wrote it.

When it is the pairing that repeats, not just the riff, put the lick in a
group with what comes before it and repeat the group (see
[Groups and repeats](#groups-and-repeats)):

```
[3B 2F# {Riff3 = G - - - - | D 4 - - - | A - 4 5 4 | E - - - -}]x4
```

Braces recall a lick, and `=` recalls a section. `{Riff1}` plays the lick
named Riff1. `=Riff1` looks for a section called Riff1, and if there is
none it prints the text `=Riff1` as it is. Braces define a lick and braces
recall it.

If you want the notes printed only once, where the lick was written, with
every recall showing just `Riff1 (x3)`, set **⚙ Layout** → Recalled licks
→ Name only (`lick_refs: "name"`; the default is `"tab"`). A reference
that finds no lick keeps its braces on paper, so a misspelt name shows. In
both cases the lick's name prints in the lick blue.

Names match loosely (`{riff1}` finds `Riff1`) and cannot start like a
chord: `A1` would read as a chord symbol on a chart line, so it is
refused. The **{ name = tab }** button fills in a free name for you
(`Riff1`, `Riff2` and so on), and in the browser every lick you have named
appears in the palette as a button that inserts a reference to it. A lick
you name while typing gets its button the next time the section card is
redrawn, after Save for instance.

Transposing the section moves a lick's frets with it: same strings,
shifted positions. A fret that would fall off either end of the neck stays
as it was, instead of being moved to a position you would not play.

## Playing techniques

A lick position, or a cell of a tab grid, can carry the usual tab letters
after the fret:

| Typed | Means |
|---|---|
| `5h7` | hammer-on from 5 to 7 |
| `7p5` | pull-off from 7 to 5 |
| `5/7` | slide up from 5 to 7 |
| `7\5` | slide down from 7 to 5 |
| `7b9` | bend at 7 up to the pitch of 9 |
| `7b9r7` | bend and release back to 7 |
| `7~` | vibrato |
| `x` | muted (dead) note |
| `5^` | accented note |

```
{G 5h7 7p5 - | D - - 7b9r7}
```

They print as typed, so a position with a technique is wider than a plain
fret, and the columns around it widen to keep the strings lined up.
Transposing moves every fret in a position (`5h7` up two is `7h9`). A
position that would go past fret 24 stays as written.

Palm muting covers a run of notes, so it is written as a run on the chart
line (see below), not as a letter on a fret.

## Marks

| Typed | Means |
|---|---|
| `\|:` | repeat-open barline |
| `:\|` | repeat-close barline |
| `\|1.` | 1st ending: a bracket marked "1." over what follows, up to the next ending or the `:\|` that closes it |
| `\|2.` | 2nd ending: a bracket marked "2." up to the next ending or the end of the section, carried over a `//` line break |
| `//` | line break: start a new line here (layout only, plays nothing) |
| `//>` | line break, indented one step per `>` |
| `%` | `simile`: play like the previous bar |
| `rest` | rest or tacet: do not play this measure |
| `coda` | coda: prints as a boxed coda sign with "Coda" beside it (`render.has_coda()` checks for it) |
| `segno` | segno mark |
| `dc` | D.C. (da capo) |
| `ds` | D.S. (dal segno) |

## Dynamics, accents and runs

How loud to play is written with the usual music words, as a word of its
own where the change happens:

```
p [pm 5A 5A 5A 5A] mf 7D^ 7D
```

| Typed | Means |
|---|---|
| `pp` `p` `mp` `mf` `f` `ff` | dynamics, printed in bold where they are typed |
| `5A^`, `Am^` | accent: a `>` over that note or chord |
| `[pm 5A 5A 5A]` | palm mute over the run: "P.M." and a dashed line to the last note |
| `[cresc A D E]` | get louder over the run: an opening hairpin (TXT: `cresc.---`) |
| `[dim E D A]` | get softer over the run: a closing hairpin (TXT: `dim.---`) |

Dynamics are lower case only, so `f` is forte and `F` stays the chord.
`pm`, `cresc` and `dim` go first inside the brackets, and a run can repeat
like any group, `[pm 5A 5A]x2`. Typed on their own, outside brackets,
they are an error that says where they go.

The marks over a line (endings, runs, accents) print on a row of their
own above the fret numbers, in the TXT, the PDF and the live preview.

## Typing it without remembering it

Each section's chart line has a palette of buttons beside it: `rest`,
`%`, `|:`, `:|`, `|1.`, `|2.`, `x2`, `[ ]x2`, `""`, `segno`, `coda`, `dc`,
`ds`. Each one puts its notation at the cursor, separated by a space from
what is there, and leaves the cursor inside the brackets or quotes when
that is where you will type next. The insert sends the same `input` event
as typing, so the live parse, the error messages and the preview behave
exactly the same, and the chart line stays text you can edit by hand.

As you type a word in the browser app, a short list offers what fits: the
song's riffs for a plain word, its sections after `=`, and its named licks
after `{`. Use the arrow keys to choose, Enter or Tab to take one, and Esc
to close the list.

The **Notation ⌘** panel explains every mark in full: what `segno`,
`coda`, `dc` and `ds` tell a player to do, and the usual
D.S., segno and coda shape they combine into.

## Sections that are not a chart line

A section's render mode can be Free instead of Chart, Tab or Both. The
section becomes one plain text box, stored as typed in `free_text`.
Nothing in it is parsed, transposed, checked or lined up, and the exports
print it exactly as typed. Use it for what this grammar does not cover.

Switching to Free leaves the section's `items` alone, so switching back
restores the chart line exactly. That is also why a reference to a Free
section uses its text and not those unused items.

## What a parse error looks like

A line that does not parse (an unmatched `[`, a bad fret, a repeat with
nothing before it, `all` after something that is not a section reference,
and so on) never clears what you typed. The error shows in place, with
the last good render still on screen above it, in both the desktop editor
bar and the browser's chart-line field.
