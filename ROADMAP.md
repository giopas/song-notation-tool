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
| **v0.18.1** | **Native app window (not frameless) + Save & Close** — `webserver.py` opens a resizable native window when `pywebview` is installed, not a browser tab; a "⏻ Save & Close" button (and new `/api/quit` route) sits alongside its own close button. See [CHANGELOG.md](CHANGELOG.md). |
| **v0.19.0** | **Web UI transpose + Lyrics panel** — Transpose (whole song or one section) now works the same in the browser as the desktop app; a new Lyrics panel (type/paste, import a `.txt` file, or a browser-search link) is available in both front ends, stored non-destructively and never auto-fetched or exported. See [CHANGELOG.md](CHANGELOG.md). |

| **v0.20.0** | **Quick-insert palette, free-text sections, `//` line breaks, Fit to one page.** See [CHANGELOG.md](CHANGELOG.md). |
| **v0.21.0** | **Licks read as tab** — printed as `\|G\|-5-7-5-\|` string rows in blue, an empty-lick button built from the section's instrument, rests in grey, and row roles/spans so both front ends colour from one source. See [CHANGELOG.md](CHANGELOG.md). |
| **v0.30.0** | **Notation and round trip** — techniques on licks and tab, dynamics, accents and runs (palm mute, cresc., dim.), ending brackets and a boxed coda in the PDF, parts for several instruments in one section, the riff library and suggestions while typing in the browser, and every PDF carries its song so it can be opened again. See [CHANGELOG.md](CHANGELOG.md). |
| **v0.29.0** | **Standalone packages** — macOS (Apple silicon and Intel), Windows (installer and zip) and Linux, built by a release workflow, with an update check and Update and restart. See [CHANGELOG.md](CHANGELOG.md). |
| **v0.28.0** | **One export writer** — every front end checks, builds and writes exports through `export.py`; the command line never replaces an existing file without `--force`; golden-file tests for ten invented songs. See [CHANGELOG.md](CHANGELOG.md). |
| **v0.27.0** | **An app icon** — a fret number over a chord, as the charts are written by hand; SVG, PNG, `.icns` and `.ico`, shown in the window, the web front end, the README and the wiki. See [CHANGELOG.md](CHANGELOG.md). |
| **v0.26.0** | **Song files named after the song** — `Artist - Title.sng`, renamed on Save when either changes, never overwriting another song; Open example song always gives a fresh example; atomic saves. See [CHANGELOG.md](CHANGELOG.md). |
| **v0.25.0** | **Change and move the songs folder** — pick a new folder from the app's window and move your songs there (only `.sng` files, nothing overwritten, old folder kept), or just point the app at it. See [CHANGELOG.md](CHANGELOG.md). |
| **v0.24.0** | **Lyric placement and a Layout menu** — six places the words can go (none; all together at the start, the end or down a left column; with each section beside or under its chart), chosen once per song; every print setting moved into ⚙ Layout at the bottom left. See [CHANGELOG.md](CHANGELOG.md). |
| **v0.23.0** | **Lyrics that carry their own structure, named licks, chord shapes** — `=== Section ===` markers in the lyric sheet (drag a section in, "Split by markers" hands each block out); a section's words print *beside* its chart rather than under it; `{Riff1 = …}` names a lick and `{Riff1}x3` recalls it; a Chords dialog prints chord-shape diagrams at either end of the chart. See [CHANGELOG.md](CHANGELOG.md). |
| **v0.22.0** | **Two-column printing** — a Columns setting (auto / 1 / 2) that splits the page when it makes the chart bigger or saves a sheet, with a rule down the gutter; fit and column count are now decided together. See [CHANGELOG.md](CHANGELOG.md). |

---

## 🎯 Planned — v0.22 and beyond

Specified in [DESIGN_v0_21.md](DESIGN_v0_21.md), in the order they are
worth building:

- [x] **Round-trip import** — done in v0.30.0: every PDF carries its
      song, and Import song… (browser), Open (desktop) and `cli.py
      extract` read it back. A TXT export cannot carry it, so only the
      PDF reopens
- [x] **Techniques on tab and licks** — done in v0.30.0, as tab letters
      (`5h7`, `7p5`, `5/7`, `7\5`, `7b9r7`, `7~`, `x`); palm muting is a
      run on the chart line, `[pm …]`
- [x] **Dynamics and accents** — done in v0.30.0 on the chart line:
      `pp`…`ff`, `5A^`, `[cresc …]`, `[dim …]`. Accents inside a lick
      use `^` after the fret
- [x] **Parts — several instruments in one section** — done in v0.30.0;
      edited in the browser app, ⚙ Layout → Instruments picks what prints
- [x] **Lyrics under the chart** — answered in v0.23.0, and answered
      *beside* rather than under: the chart keeps a narrow left column and
      the words run down their own to the right of it, with no attempt to
      align chord to syllable

---

## 🔜 Next up (v0.18.1+)

Deferred out of v0.17 that v0.18 didn't touch, plus what v0.18 itself
opens up:

- [x] **Graphical ending brackets and a boxed coda block** — done in
      v0.30.0, in the PDF, the TXT and the browser preview
- [x] **Autocomplete in the chart editor bar** — done in v0.30.0 in the
      browser app: riffs, `=` sections and `{` named licks. The desktop
      editor bar does not have it yet
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
- [x] **Web UI: riff library management** — done in v0.30.0: the
      Riffs 🔁 dialog and **→ riff**. Duplicating a riff as a reference is
      still desktop-only
- [ ] **Desktop: edit parts** — the desktop app lists a section's parts
      and keeps them, but they are edited in the browser app
- [ ] **`webserver.py` authentication** — currently no auth at all,
      fine on `localhost` or a trusted home network; worth a lightweight
      token before recommending `--host 0.0.0.0` on anything else

## 💡 Other ideas

- [ ] **MIDI playback** — hear the tab back at a set tempo
- [ ] **Chord diagrams** — visual fretboard popup for common chord shapes
- [ ] **Time signature support** — 3/4, 6/8, etc.
- [ ] **Import from Guitar Pro / GuitarPro format** — stretch goal
- [ ] **Undo / Redo stack** — per section
- [x] **Multiple instruments per song** — parts, v0.30.0
- [ ] **Auto-save / crash recovery**
- [ ] **Dark/light theme preference saved between sessions**
- [ ] **Keyboard shortcuts reference card** (in-app)

---

Have an idea? Open a [Feature Request](../../issues/new?template=feature_request.md)!
