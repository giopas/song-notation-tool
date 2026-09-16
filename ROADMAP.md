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
| **v0.17** | **Song map + chart editor bar** — the UI rewrite: whole-song map, live-parsing chart editor, riff library with usage tracking, reordering, non-destructive transpose end to end, Stage View, page-break-safe PDF. See [CHANGELOG.md](CHANGELOG.md) and [RELEASE_NOTES_v0.17.0.md](RELEASE_NOTES_v0.17.0.md). |

---

## 🔜 Next up (v0.17.1+)

Deferred out of v0.17 — see `RELEASE_NOTES_v0.17.0.md` for the full list
and why:

- [ ] **Graphical ending brackets and a boxed coda block** — `render.py`
      already resolves the spans (`mark_spans()`, `has_coda()`); nothing
      draws them yet beyond inline text tokens
- [ ] **Autocomplete in the chart editor bar** — `=` for section names,
      bare identifiers for riff names
- [ ] **Tab-grid measure copy/paste** and a per-measure beats picker
      (both existed pre-v0.17; dropped in the rewrite for time)
- [ ] **§11's architecture fork** — Tkinter vs. Flask + browser SPA,
      still open; the pure-data layers (`model`/`grammar`/`transpose`/
      `render`) don't care either way
- [ ] **Real one-page-in-five-minutes measurement** — the page estimate
      and Enter/Tab-driven entry are built toward the target, not yet
      benchmarked against an actual song

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
