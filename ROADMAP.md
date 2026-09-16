# Roadmap — Song Notation Tool

> This is a personal / vibecoded project. The roadmap is a wish list, not a commitment. Items may be added, removed, or reprioritised at any time.

The organizing goal is a typed version of the one-page stage chart the
author writes by hand — see `DESIGN_v0_16.md` section 1 for the full
vocabulary that implies. Everything below is ordered by how directly it
serves that, not by difficulty.

---

## ✅ Released

| Version | Highlights |
|---|---|
| v0.1 | Initial release — sections, layers (tab/chords/notes/lyrics), export TXT |
| v0.2 | Light/dark theme engine, ToolTips, ttk styling, layer toggle buttons |
| v0.3 | Artist name field, section copy dialog, APP_VERSION constant |
| v0.4 | Two-row topbar, Drop D toggle, Transpose dialog |
| v0.5 | Fixed macOS button colours; horizontal tab Entry per string with dashes |
| v0.6 | Fixed macOS layer-toggle colours; custom scroll indicators; copy/paste fix |
| v0.7 | Section linking — propagate edits across linked sections |
| v0.8 | Resizable left panel sash; measure grid wrap mode |
| v0.10 | TXT/PDF export fixes; Artist before Title; cleaned-up chord/note display |
| v0.11 | Chromatic note list with enharmonic aliases; transpose all layers; root-note picker |
| v0.12 | Variable tab beats (8/16/32); protected dashes; smart transposition |
| v0.13 | Beats selector in editor toolbar; hamburger menu for narrow windows |
| v0.14 | Minor stability and layout fixes |
| v0.15 | Ongoing refinements before this changelog existed |
| **v0.16** | **Foundation release** — new data model, chart grammar, non-destructive render-time transposition, compact exports, v0.15 migration, repo hygiene. See [CHANGELOG.md](CHANGELOG.md). |

---

## 🔜 Toward the one-page chart (v0.17+)

These are the items that directly close the gap between "the app can
export a song" and "the export looks like the reference sheets":

- [ ] **Chart row inline editor with live re-render** — the v0.16 chart
      field validates on confirm; it should show the rendered two-line
      chart (fret over symbol) as you type
- [ ] **Song map whole-song view** — see every section's chart line on one
      screen, the way the handwritten page shows the whole song at once
- [ ] **Riff library panel** — create, name, and reuse `block`s from the
      UI (the model and grammar already support `block_ref`; there's no
      editor for `blocks` yet)
- [ ] **Mark and ending rendering** — `|:`, `:|`, 1st/2nd endings, segno,
      coda, `simile` (`%`) parse today (`grammar.py`) but don't render in
      TXT/PDF output yet
- [ ] **One-page pagination** — compact mode drops empty strings and
      collapses references, but doesn't yet guarantee a fit; needs real
      pagination logic against the Portrait A4 target
- [ ] **Stage/print view** — a distraction-free, large-type render meant to
      be read from a music stand, not edited

## 💡 Other ideas

- [ ] **MIDI playback** — hear the tab back at a set tempo
- [ ] **Chord diagrams** — visual fretboard popup for common chord shapes
- [ ] **Time signature support** — 3/4, 6/8, etc.
- [ ] **Import from Guitar Pro / GuitarPro format** — stretch goal
- [ ] **Undo / Redo stack** — per section
- [ ] **Multiple instruments per song** — e.g. guitar + bass in the same project
- [ ] **Auto-save / crash recovery**
- [ ] **Dark/light theme preference saved between sessions**
- [ ] **Keyboard shortcuts reference card** (in-app)

---

Have an idea? Open a [Feature Request](../../issues/new?template=feature_request.md)!
