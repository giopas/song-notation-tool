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
| **v0.17** | **Song map + chart editor bar** — the UI rewrite: whole-song map, live-parsing chart editor, riff library with usage tracking, reordering, non-destructive transpose end to end, Stage View, page-break-safe PDF. See [CHANGELOG.md](CHANGELOG.md). |
| **v0.18** | **Headless & browser use, plus a UX pass** — `export.py` extracted to a pure TXT/PDF engine; `cli.py` (convert/batch/lint, no window); `webserver.py` + `web/` (stdlib-only local server and browser front end); a "Start here" panel, help strip, and live Preview window in the desktop app. Resolves the "§11 architecture fork" item below. See [CHANGELOG.md](CHANGELOG.md). |

---

## 🔜 Next up (v0.18.1+)

Deferred out of v0.17 that v0.18 didn't touch, plus what v0.18 itself
opens up:

- [ ] **Graphical ending brackets and a boxed coda block** — `render.py`
      already resolves the spans (`mark_spans()`, `has_coda()`); nothing
      draws them yet beyond inline text tokens
- [ ] **Autocomplete in the chart editor bar** — `=` for section names,
      bare identifiers for riff names
- [ ] **Tab-grid measure copy/paste** and a per-measure beats picker
      (both existed pre-v0.17; dropped in the rewrite for time)
- [ ] **Real one-page-in-five-minutes measurement** — the page estimate
      and Enter/Tab-driven entry are built toward the target, not yet
      benchmarked against an actual song
- [ ] **Web UI: tab-grid (measure) editing** — the browser front end
      shipped in v0.18 edits `render:"chart"` sections fully; a
      `render:"tab"`/`"both"` section's measure grid is read-only there
      for now (ships read-only/preview first, per the phased roadmap;
      full grid editing is the natural v0.18.1 follow-up)
- [ ] **Web UI: riff library management** — creating/renaming/deleting
      riffs and Promote-to-riff/Duplicate-as-reference are desktop-only
      for now; the web editor can *use* an existing riff reference but
      not manage the library
- [ ] **`webserver.py` authentication** — currently no auth at all,
      fine on `localhost` or a trusted home network; worth a lightweight
      token before recommending `--host 0.0.0.0` on anything else
- [ ] **Lyric importer** — local-file and paste import of lyrics to line
      up against notation faster; a web-search-assisted version should
      stay a human-reviewed assist (search results shown, user pastes),
      never an automatic fetch-and-store, for both copyright and
      section-boundary-accuracy reasons

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
