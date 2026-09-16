# Song Notation Tool — v0.17 design

**The new surface.** Song map view, riff library, inline chart editing, one-page print output. Depends entirely on the model shipped in [v0.16](DESIGN_v0_16.md) — read that first.

---

## 1. Goal

v0.16 makes the compact chart *representable*. v0.17 makes it *fast to write*.

The measure of success is blunt: **a 12-section song should take under five minutes to enter, and fit on one page.** The handwritten reference sheets take about that long with a pen. Anything slower than the pen is a failed tool.

---

## 2. What is wrong with the v0.15 editor

| Problem | Consequence |
|---|---|
| One section visible at a time | The whole song is never on screen, yet every reference sheet is a whole-song view. Structure is the thing being written, and the editor hides it. |
| Entry is cell-by-cell | 4–6 string rows × 8 beats × 8 measures per section. Hundreds of keystrokes and mouse moves for a section that is one line on paper. |
| Section reordering is two buttons | Move-up / move-down on a listbox. Reordering a 12-section song is dozens of clicks. |
| Sections are named, then configured in a dialog | Every structural change routes through a modal. |
| Repeats are invisible | `repeat` is a field in a dialog, shown as `[4m]` in the listbox. On paper it is the most prominent thing on the line. |

---

## 3. Song map

The main view becomes a vertical list of section rows. This replaces the listbox + single-section editor split entirely.

```
  ┌──────────────────────────────────────────────────────────┐
  │  Sample Song · Sample Band · key B · 4/4        [1 page]    │
  ├──────────┬───────────────────────────────────────────────┤
  │   Intro  │  7   9  10   7  │ 7   9 │×2                    │
  │          │  B  F#   G   E  │ B  F# │                      │
  │ Verse 1  │  riff 1  ×4            1 riff = 4 bars         │
  │ Chorus 1 │  riff 2  ×4                                    │
  │ Verse 2  │  = verse 1  ×4                                 │
  │  Bridge  │  ▼ tab · 4 bars · bass                         │
  │          │    A|  -  7  -  -  7  -  -  -                  │
  │          │    E|  -  -  10 -  -  8  -  -                  │
  │   Outro  │  coda   = chorus 1  ×1 all                     │
  ├──────────┴───────────────────────────────────────────────┤
  │  chorus 1 ›  riff2 x4▏                                    │
  ├──────────────────────────────────────────────────────────┤
  │  Riffs   [riff 1 · 4 bars]  [riff 2 · 4 bars]  [+ new]    │
  └──────────────────────────────────────────────────────────┘
```

### 3.1 Layout

- **Left gutter** — section name, right-aligned, fixed width. Click to rename inline. No dialog.
- **Row body** — the rendered chart line for that section. Tokens are stacked pairs: fret above in muted small type, symbol below at normal size. This stacking is what keeps the position unambiguous against complex chords like `F#m7`.
- **Section repeat** — rendered at the end of the row as `×4`, in the same visual weight as on paper.
- **Annotation** — right-aligned in the row, muted, e.g. "1 riff = 4 bars".
- **Tab sections** — collapsed to a one-line summary (`▼ tab · 4 bars · bass`) that expands in place. A song with one notated bar shows one expanded row and eleven collapsed ones.
- **Page indicator** in the header — live count of how many printed pages the document currently occupies. The one-page target should be visible while writing, not discovered at export.

### 3.2 Interaction

- **Drag a row by its gutter** to reorder. Replaces move-up / move-down.
- **Click a row body** to focus it in the editor bar (§4).
- **Enter** on a focused row creates a new section below it.
- **Alt+↑ / Alt+↓** moves the focused section without the mouse.
- Rows render read-only until focused. No always-live input widgets — this is what makes a 20-section song scroll smoothly in Tk.

---

## 4. Chart editor bar

A single text field, pinned below the map, showing the grammar source for the focused section. Typing re-renders that row live.

```
  chorus 1 ›  riff2 x4▏
```

- The line is generated from `items` by the v0.16 unparser on focus, and parsed back on every keystroke (debounced ~120 ms).
- **Parse errors do not clear the row.** The last valid render stays on screen; the error is shown inline under the field with the character offset. Never destroy work on a half-typed token.
- **Tab** commits and moves to the next section. **Escape** reverts to the last valid state.
- Autocomplete on `=` (section names) and on bare identifiers (block names), since both must match existing ids.

This bar is the entire editing surface for chart sections. The tab grid remains, unchanged from v0.15, for `render: tab` sections only.

---

## 5. Riff library

A strip along the bottom of the map, plus a panel when expanded.

- Each riff is a chip: name, bar count, and a usage badge showing how many sections reference it.
- **Click a chip** to open the riff in the same editor surface as a section — a riff is structurally identical to a section, so it reuses the row renderer and the editor bar.
- **Editing a riff updates every reference at once.** This is the single largest labour saving in the tool and should be demonstrated in the README.
- **Deleting a referenced riff** is blocked, with a list of the referencing sections. Never leave dangling refs.
- **Promote to riff** — select a run of items in a section and turn it into a named block, replacing the selection with a `block_ref`. This is how riffs get created in practice: you write it once inline, then notice it repeats.

Per-reference transpose (`riff1 x3 +2`) is set in the grammar, not in the panel. The panel shows resolved offsets read-only, so it is obvious when two references differ.

---

## 6. Marks, endings and the coda box

The `mark` item kind ships in v0.16 as data. v0.17 renders it.

| Mark | Rendering |
|---|---|
| `repeat_open` / `repeat_close` | `‖:` and `:‖` barline glyphs inline |
| `ending_1` / `ending_2` | A bracket above the run, numbered, spanning until the next mark |
| `coda` | A boxed block, visually separated as on the Tonight Tonight sheet |
| `segno` / `dc` / `ds` | Inline symbol with a tooltip spelling out the navigation |
| `simile` (`%`) | Inline repeat-previous-bar glyph |

Endings need a span, not a point. Render `ending_1` as opening a bracket that closes at the next `ending_*` or end of section. The data stays a flat sequence; only the renderer resolves spans.

---

## 7. Print and PDF

v0.16 makes the output compact. v0.17 makes it deliberate.

- **Portrait A4, one page, as a target the layout actively pursues** — shrink the inter-section gap before shrinking type, and shrink type only within a floor (9 pt minimum; this is read on stage in bad light).
- **Never split a section across pages.** If it does not fit, break before it.
- **Stage view** — a read-only full-window mode: large type, high contrast, no chrome, arrow keys to scroll. The same renderer as the map with a different size scale.
- **Print stylesheet parity** — the PDF and the stage view must come from one renderer with two scales, not two code paths. The v0.15 divergence between TXT and PDF happened because they were written twice.

---

## 8. Keyboard model

The whole song should be enterable without the mouse.

| Key | Action |
|---|---|
| `Enter` | New section below |
| `Tab` / `Shift+Tab` | Next / previous section |
| `Alt+↑` / `Alt+↓` | Move section |
| `Escape` | Revert current edit |
| `Ctrl/Cmd+R` | Promote selection to riff |
| `Ctrl/Cmd+D` | Duplicate section as a reference (`=`) rather than a copy |
| `Ctrl/Cmd+P` | Stage view |

`Ctrl+D` producing a reference rather than a copy is deliberate. Copies are what produced the duplicate Verse 2 / Chorus 2 pages in the v0.15 Sample Song output.

---

## 9. Build order

1. Row renderer — `items` → widgets. Read-only. Shared with print at a different scale.
2. Song map list with focus, drag reorder, inline rename.
3. Editor bar with live parse, error display, and revert.
4. Riff strip, then the panel, then promote-to-riff.
5. Marks and ending spans in the renderer.
6. Pagination and stage view.
7. Retire the v0.15 single-section editor.

Steps 1 and 3 carry the risk. Everything else is assembly.

---

## 10. Acceptance criteria

- [ ] A 12-section song is enterable in under five minutes, keyboard only
- [ ] The Sample Song reference case renders on one portrait A4 page
- [ ] Editing `riff 1` visibly updates all four referencing sections
- [ ] A half-typed token never clears the rendered row
- [ ] Deleting a referenced riff is blocked and names the referencing sections
- [ ] Drag-reordering 12 sections requires no dialog
- [ ] The stage view and the PDF are produced by the same renderer
- [ ] No section is ever split across a page break

---

## 11. Open question — the architecture fork

The project's development reference is `giopas/qlc-plus-swiss-knife-tool-script`, which has since moved from single-file Tkinter to **Flask + browser SPA** (`app.py`, `core/` modules, `routes/` blueprints, `static/css/tokens.css`, a sidebar shell). Song Notation Tool is still a 2,978-line single-file Tkinter app.

Most of what v0.17 asks for — stacked-pair tokens, drag reorder, bracket spans, a boxed coda, print pagination — is routine in HTML/CSS and awkward in Tk. PDF export via a browser print stylesheet would also replace the hand-rolled PDF writer entirely.

This document is deliberately written to be **UI-toolkit agnostic**. The v0.16 layers (`model`, `grammar`, `transpose`, `render`) have no Tk dependency and port unchanged either way. The decision can therefore be taken at the start of v0.17 implementation rather than now — but it should be taken **before** step 1 of §9, because the row renderer is the piece that would otherwise be written twice.
