# Song Notation Tool — v0.16 design

**Foundation release.** New data model, chart grammar, render-time transposition, compact exports, repo hygiene. No UI rewrite — that is [v0.17](DESIGN_v0_17.md).

---

## 1. What this project is actually for

The author writes bass and guitar parts by hand, on one sheet of A4, to read on stage. Those sheets are the target output. They share a small, consistent vocabulary:

- A **fret number written above a note or chord symbol** — `5` over `A` means "the A played at the 5th fret". This is the atomic unit on most sheets.
- **Repeat counts at every level** — per group `(x2)`, per section `(x4)`, and the shorthand `(x1 all)`.
- **Named riffs defined once and referenced by name** — `RIFF 1 (x3)`, with a legend such as "1 riff = 4 bars".
- **Back-references** — `VERSO (x1)` meaning "same as the verse above".
- **One small stave for the one bar that needs it** — never the whole song.
- **Free annotations** in plain language — "stacco sul secondo".
- **Navigation marks** — 1st/2nd endings, repeat barlines, a boxed coda.

v0.15 supports none of this directly. It supports exactly one thing: a full measure-by-measure tab grid.

---

## 2. Honest audit of v0.15

Reference case: `Sample Band — Sample Song`, a 6-section bass song, exported by v0.15.

| Symptom | Detail |
|---|---|
| Output is 3 PDF pages | The same song on paper is ~12 lines |
| ~85% of the output is dashes | G and D strings are empty in every measure of every section, and printed anyway |
| Sections duplicated in full | Verse 2 is identical to Verse 1, Chorus 2 to Chorus 1 — each written out twice |
| Bad token accepted | Intro M4, E string contains `01` — a typo the beat-cell entry never rejected |
| TXT and PDF disagree | `Interlude 2` (note names, no tab) renders as a header with notes in TXT, and as empty measure columns in PDF |
| TXT lines overflow | `_build_song_lines` defines `W = 80` and then emits ~200-character rows |
| Measure wrapping differs | PDF wraps 8 measures to 6 + 2; TXT puts all 8 on one line |
| Transposition is lossy | `_transpose_fret_token` rewrites cells in place and clamps at the fret floor; +5 then −5 does not restore the original |

None of these are isolated bugs. They all follow from the same root cause: **the tab grid is mandatory and the compact chart does not exist.**

---

## 3. The model

A section is a flat sequence of items, and an item can be a group of items. That single recursive shape covers every mark on the reference sheets.

### 3.1 Document schema (`.sng`, format 2)

```json
{
  "format": 2,
  "app_version": "0.16",
  "meta": {
    "title": "Sample Song",
    "artist": "Sample Band",
    "key": "B",
    "time": "4/4",
    "bpm": ""
  },
  "transpose": 0,
  "blocks": {
    "riff1": {
      "id": "riff1",
      "name": "Riff 1",
      "bars": 4,
      "beats_per_bar": 8,
      "instrument": "Bass (4-string)",
      "items": []
    }
  },
  "sections": [
    {
      "id": "sec1",
      "name": "Verse 1",
      "type": "Verse",
      "instrument": "Bass (4-string)",
      "repeat": 4,
      "transpose": 0,
      "render": "chart",
      "annotation": "",
      "items": []
    }
  ]
}
```

- `render` is one of `chart`, `tab`, `both` — per section. This is the switch that lets one document hold a one-line section map and a fully notated bar.
- `blocks` is the riff library. Keyed by id.
- `transpose` appears at three levels (document, section, block reference). See §5.

### 3.2 Item kinds

Every item is `{"kind": ..., ...}`. Six kinds:

**`token`** — a note or chord, with an optional fret position.

```json
{"kind": "token", "symbol": "F#m7", "fret": 2}
```

`fret` is optional and may be `null`. The fret refers to **the note itself**, not to a string — see §4.3. `symbol` is the full chord or note symbol as written (`E`, `F#`, `Bb`, `Am7`, `C/G`, `A5`).

**`group`** — a bracketed run with its own repeat.

```json
{"kind": "group", "repeat": 2, "items": [ ... ]}
```

Groups nest.

**`block_ref`** — a reference to a named riff.

```json
{"kind": "block_ref", "block": "riff1", "repeat": 3, "transpose": 0}
```

**`section_ref`** — "as before".

```json
{"kind": "section_ref", "section": "sec1", "repeat": 1, "all": true, "transpose": 0}
```

`all` renders the `x1 all` shorthand.

**`measure`** — one bar of full tablature. This is the v0.15 tab cell, unchanged.

```json
{"kind": "measure", "beats": 8, "strings": {"G": "- - - - - - - -", "E": "- 7 - - - - - -"}}
```

**`mark`** — a navigation symbol.

```json
{"kind": "mark", "mark": "coda"}
```

`mark` ∈ `repeat_open` `repeat_close` `ending_1` `ending_2` `segno` `coda` `dc` `ds` `simile`

**Annotations** are not items. They are the section-level `annotation` string, plus an optional per-group `note` field for inline remarks like "stacco sul secondo".

### 3.3 What this replaces

`link_id` is removed. Linked sections become either a `section_ref` (display-level: "as Verse 1") or a shared `block` (data-level: same riff used in two places). Both are strictly more expressive than the current coupling, and `section_ref` is what removes the duplicate Verse 2 / Chorus 2 pages from the Sample Song export.

---

## 4. Chart grammar

The chart row is **one text field per section**. The user types a line, the app parses it to `items`, and renders. This is the whole point of the release: one line of typing per section rather than 4–6 string rows × 8 measures of cell entry.

### 4.1 Grammar

```ebnf
line        ::= item*
item        ::= token | group | block_ref | section_ref | mark | annotation

token       ::= fret? symbol
fret        ::= digit digit?
symbol      ::= note accidental? quality?
note        ::= "A".."G"
accidental  ::= "#" | "b"
quality     ::= ( letter | digit | "/" | "+" | "-" | "(" | ")" )*

group       ::= "[" item* "]" repeat?
block_ref   ::= ident repeat? shift?
section_ref ::= "=" ident repeat? shift?
repeat      ::= "x" digit+ "all"?
shift       ::= ("+" | "-") digit+
mark        ::= "|:" | ":|" | "|1." | "|2." | "%" | "coda" | "segno" | "DC" | "DS"
annotation  ::= '"' text '"'
ident       ::= letter ( letter | digit | "_" )*
```

Items are whitespace-separated. `repeat` and `shift` bind to the item immediately preceding them.

### 4.2 Examples

| Typed | Means |
|---|---|
| `0E` | E played open |
| `7E` | E played at the 7th fret |
| `5A 5D 5A 8F 8C 5A 5D` | seven positioned tokens |
| `C` | C with no position given |
| `2F#m7` | F#m7 at the 2nd fret |
| `[5A 5D]x2` | bracketed group, played twice |
| `riff1 x3` | riff 1, three times |
| `riff1 x3 +2` | riff 1, three times, a tone up |
| `=verse1 x1 all` | as verse 1, once, all of it |
| `\|1. 5A \|2. 3C` | first and second endings |
| `"stacco sul secondo"` | annotation |

### 4.3 Resolved ambiguity: the fret must lead

An earlier proposal allowed both `5A` and `A5`. **That is wrong and must not be implemented.** `A5` is a legitimate power-chord symbol. The fret is therefore always a leading number and never trailing:

- `5A` → fret 5, symbol `A`
- `A5` → no fret, symbol `A5` (A power chord)
- `5A5` → fret 5, symbol `A5`

This is also why the fret is stored as separate metadata and never concatenated into `symbol`. Rendering it stacked above the symbol keeps it unambiguous against complex chords.

### 4.4 Resolved ambiguity: block names

A bare identifier is read as a `block_ref`. A block whose name parses as a token would be ambiguous (`A1`, `b2`). **Validation rule:** reject block names matching `^\d{0,2}[A-Ga-g]([#b])?` at creation time. The UI must show the error inline, not silently rename.

### 4.5 Round-trip requirement

`items` is the single source of truth in the saved file. The editor line is regenerated from `items` by an unparser. There is no second stored copy of the typed text, so no divergence is possible.

**Acceptance test:** `parse(unparse(items)) == items` for every fixture in `tests/fixtures/`.

---

## 5. Transposition

### 5.1 Three additive scopes, applied at render time

```
effective = document.transpose + section.transpose + ref.transpose
```

A token inside `riff1`, referenced from Chorus 2, resolves with the document offset, the section's offset, and that reference's offset. **The block definition is never rewritten.** A riff referenced four times transposes once per token at render; rewriting the definition would double-count against the per-reference offsets and silently corrupt the other three references.

### 5.2 Token rule

The rule that preserves fingering: **note + n semitones, fret + n.** Same string, slide up.

```
0/E  +1  →  1/F
7/E  +1  →  8/F
```

### 5.3 Resolution

```python
MAX_FRET = 24   # already defined in v0.15

def resolve(token, eff):
    symbol = transpose_symbol(token.symbol, eff)
    if token.fret is None:
        return Resolved(symbol, None, flag=None)
    fret = token.fret + eff
    if fret < 0:
        fret += 12
        return Resolved(symbol, fret, flag="octave_up")
    if fret > MAX_FRET:
        fret -= 12
        return Resolved(symbol, fret, flag="octave_down")
    return Resolved(symbol, fret, flag=None)
```

If the corrected fret is still out of range, render the symbol with no fret and flag `unplayable`. Flags are a **display state**, never written to the document — a token unplayable at +7 becomes playable again at +2 with nothing lost in between.

### 5.4 Non-destructive, with an explicit bake

Transposition is applied when drawing, exporting and printing. It never rewrites stored notes. This makes "what does this sound like in D?" a free, reversible question.

Add a **Bake transpose** action that writes the offsets into the tokens and resets all three counters to zero. Destructive, and only on request.

### 5.5 Key metadata follows

`meta.key` is displayed with `document.transpose` applied. Transposing +2 from B shows C#. The stored value does not change until bake.

---

## 6. Migration from v0.15

Old files must load without loss. `format` is absent in v0.15 files — treat a missing `format` as 1.

| v0.15 | v0.16 |
|---|---|
| `layers.tab[m]` (dict of string rows) | `{"kind": "measure", "beats": ..., "strings": ...}` |
| `layers.chords[m]` | `{"kind": "token", "symbol": ..., "fret": null}` |
| `layers.notes[m]` | `{"kind": "token", "symbol": ..., "fret": null}` — merged into chords if chords is empty for that measure, otherwise kept as a parallel row |
| `layers.lyrics[m]` | section `lyrics` array, preserved verbatim |
| `measure_beats` | `beats` on each `measure` item |
| `tab_beats` | section default, retained |
| `link_id` | `section_ref` on the later of the two linked sections |
| `visible` | `render`: `tab` if tab was visible, else `chart` |
| `repeat` | unchanged |

Every migrated section gets `render: "tab"` if it has any non-empty tab measure, so old documents look the same on first open.

**Acceptance test:** load each fixture in `tests/fixtures/v015/`, save, reload — no data loss, and the TXT export of the migrated file is content-equivalent to the v0.15 export of the original.

---

## 7. Export changes

### 7.1 Shared predicate

Both exporters must use one helper:

```python
def section_has_tab(section) -> bool
```

The TXT/PDF disagreement on `Interlude 2` exists because `_build_song_lines` and `_build_pdf` each compute this inline and differ. One function, both callers.

### 7.2 Chart rendering

A chart row is two text lines, column-aligned:

```
       0    7    5    8
Verse  E    E    A    F        (x4)
```

Fret line above, symbol line below, one column per item, blank where no fret is given. Groups get bracket characters and a trailing `(xN)`. Block and section references render as their name plus repeat, never expanded.

### 7.3 Compact mode (default on)

- **Drop empty strings.** A string that is empty across the whole document, per instrument, is not printed. This alone removes the G and D rows from Sample Song.
- **Do not expand references.** `= Verse 1 (x4)` stays one line.
- **Target one page.** Portrait A4 default for PDF (v0.15 defaults to landscape). Full pagination work is v0.17.

### 7.4 TXT wrapping

`W` is currently defined and ignored. Make it a real setting — default 100 columns — and compute measures-per-line from it, matching the logic `mpl_for()` already uses in the PDF builder so both wrap identically.

---

## 8. Bug fixes in scope

1. **Fret token validation.** Reject leading zeros (`01`), non-numeric junk, and frets > `MAX_FRET`. Applies to both the chart grammar and the tab cell entry.
2. **`ENHARMONIC` has a duplicate key.** `"Bb"` is defined twice in the dict literal. Lowercase variants for `db`, `eb`, `gb`, `ab` are missing while `bb` is present.
3. **TXT `W = 80` unused** — see §7.4.
4. **TXT/PDF `has_tab` mismatch** — see §7.1.
5. **Measure wrapping differs between exporters** — see §7.4.
6. **Destructive, clamping transposition** — replaced by §5.

---

## 9. Repo alignment

The GitHub repo has no tags, no releases, no changelog, and an empty description. Releases are frequent, so this needs to exist before v0.16 ships.

- [ ] Rename `song_writer_v.0.15.py` → **`song_writer.py`**. Version lives in `APP_VERSION` and on the release asset. The version-in-filename breaks every link and script on each release.
- [ ] Add **`CHANGELOG.md`** in Keep a Changelog format, backfilled from `ROADMAP.md`'s released table.
- [ ] Tag **`v0.16.0`** and cut the first GitHub Release, with **`RELEASE_NOTES_v0.16.0.md`** (matching the `qlc-plus-swiss-knife-tool-script` convention).
- [ ] **Rewrite `README.md` goal-first.** Lead with what the tool produces — a one-page stage chart — not a feature list. Include a before/after image: a handwritten reference sheet beside the tool's output. Add a screenshots table.
- [ ] Add **`examples/`** with `sample_song.sng`, `sample_song.txt`, `sample_song.pdf`. This explains the project faster than any prose.
- [ ] Set the **repo description and topics** on GitHub (both currently empty).
- [ ] **Restructure `ROADMAP.md`** around reaching the handwritten-sheet output. MIDI playback and Guitar Pro import are currently listed while chart mode is not.
- [ ] Add **`DEVELOPMENT.md`** describing the AI-assisted release workflow, so each session starts from a consistent brief.

---

## 10. Build order

1. `model.py` — schema, item kinds, (de)serialisation, v0.15 migration. Pure data, no Tk.
2. `grammar.py` — parser and unparser. Pure functions, fully unit-testable.
3. `transpose.py` — `resolve()`, symbol transposition, fret edges. Reuses the existing `CHROMATIC`/`ENHARMONIC` tables with the §8.2 fix.
4. `render.py` — items → text lines, shared by both exporters.
5. Wire into the existing exporters; delete the duplicated inline logic.
6. Minimal UI: a single chart text field per section alongside the existing tab grid. Not the song map — just enough to exercise the grammar.
7. Repo cleanup and release.

Steps 1–4 have no Tkinter dependency. They port unchanged if the app later moves to a Flask/SPA architecture, which is why they come first.

---

## 11. Acceptance criteria

- [ ] `parse(unparse(items)) == items` for all fixtures
- [ ] All v0.15 fixtures load, save and reload without data loss
- [ ] Transposing +n then −n returns the document to its exact original state
- [ ] Sample Song exports to **one portrait A4 page** with the G and D strings absent and Verse 2 / Chorus 2 shown as references
- [ ] TXT and PDF agree on which sections have tab and how measures wrap
- [ ] `01` is rejected at entry in both the grammar and the tab grid
- [ ] `A5` parses as an A power chord with no fret; `5A` parses as A at fret 5
- [ ] Block named `A1` is rejected with an inline error

---

## 12. Out of scope

Deferred to v0.17: the song map whole-song view, the riff library panel, the chart row inline editor with live re-render, mark and ending rendering, one-page pagination, and the stage/print view. Deferred indefinitely: MIDI playback, chord diagrams, Guitar Pro import.
