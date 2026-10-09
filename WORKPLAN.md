# Work plan

The living plan for Song Notation Tool: where we are, what is left, and the
decisions in force. A new session reads this first, then `DEVELOPMENT.md`
for how the work is done. After a change here, the same file is copied to
the Claude project as `claude/WORKPLAN.md`.

## Where we are

- Released: v0.29.1, published on 9 October 2026 with all eight files
  (macOS Apple silicon and Intel, `.dmg` and `.tar.gz` each; Windows
  installer and `.zip`; Linux `.tar.gz`; `SHA256SUMS`). The release
  workflow built and smoke-tested all four systems and tested the Windows
  installer.
- On 10 October giopas installed v0.29.0 on the Mac and used **Update and
  restart** to go to v0.29.1: it worked the first time.
- v0.29.0 is published too, without Windows files (its Windows build
  stopped in the tests; fixed in v0.29.1).
- Every release from v0.26.3 on has its GitHub release.
- Items 1 to 8 of the first plan are done (see "Done" below).
- v0.30.0 is built on the branch `feat/v0.30.0`, committed and not yet
  merged or tagged: seven roadmap items in one release (see "Second plan"
  below). Waiting for giopas to push it and run the release workflow.
- Test suite: 479 tests, all passing (`python3 -m pytest -q`).

## What is left

Housekeeping, small:

1. Done: the newer `release.yml` (actions on Node.js 24) is in
   `.github/workflows/` (commit `14fd35d`), and v0.30.0 is its first
   release.
2. The three rewritten issue templates (from v0.27.1) are still not in
   `.github/ISSUE_TEMPLATE/`: commit `14fd35d` names them but only changed
   `release.yml`, and the files there are the originals from v0.15. Copy
   them in and commit them.

After v0.30.0, the next piece of work is giopas's choice; see
"Candidates for the next version" below.

### Second plan: v0.30.0 (10 October 2026)

giopas picked seven candidates and one release for all of them, with
these choices: standard tab letters for techniques, music words and `^`
for dynamics and accents, extra parts per section.

| Item | What |
|---|---|
| 1 | Round trip: every PDF carries its song (`song.sng` attached). **Import song…** in the browser, **Open** in the desktop app, `cli.py extract`. TXT cannot carry it |
| 2 | Riff library in the browser: **Riffs 🔁**, **→ riff**, `POST /api/riffs` |
| 3 | Techniques on licks and tab cells: `5h7 7p5 5/7 7\5 7b9 7b9r7 7~ x` |
| 4 | Parts: **+ Part** per section, **⚙ Layout → Instruments** (`print_instruments`). Edited in the browser only |
| 5 | Dynamics `pp`…`ff`, accent `^`, runs `[pm …]`, `[cresc …]`, `[dim …]` |
| 6 | Ending brackets and a boxed coda in the PDF, TXT and preview |
| 7 | Suggestions while typing a chart line (browser) |

Checked by Claude: the two new golden PDFs looked at page by page, the
browser app driven with Playwright (parts, riffs dialog, → riff,
suggestions, Import of a PDF), and the desktop app opening a PDF under
Xvfb. Not yet seen by giopas.

### Done (first plan, 8 to 10 October 2026)

| Item | Version | What |
|---|---|---|
| 1 | - | GitHub releases for every version from v0.26.3 |
| 2 | v0.27.1 | Plain-language pass over the README, wiki and app texts (issue templates pending, see above) |
| 3 | v0.28.0 | One export writer: `export.export_to_file()` / `prepare_export()` |
| 4 | v0.28.0 | Exports never overwrite; the CLI has `--force` |
| 5 | v0.28.0 | `export.check_document()` before every export |
| 6 | v0.28.0 | Golden files: ten invented songs in `tests/golden/` |
| 7 | v0.27.0 | App icon (option C) |
| 8 | v0.29.0, v0.29.1 | Packages for macOS, Windows and Linux, release workflow, update check and **Update and restart** |

### Candidates for the next version

From `ROADMAP.md` and the open notes, closest to the goal first. None is
started until giopas picks it.

- Editing parts in the desktop app, and suggestions in its editor bar.
- Duplicating a riff as a reference in the browser app.
- The benchmark: one real song rebuilt with the tool only and timed
  ("one page in five minutes"), with what got in the way written down.
- A token for `webserver.py --host 0.0.0.0` before it is used beyond a
  trusted network.
- A test file for the Tkinter desktop app (`tests/test_song_writer.py`).

Manual checks waiting for giopas (each stays here until reported back):

- v0.26.3 repeat bracket in a printed PDF and in the web preview.
- v0.27.0: the icon in the desktop app's window and Dock, in the native
  window, and as the favicon in a browser tab.
- v0.27.1: a read-through of the README and the wiki, and a look at the
  app's help texts, to confirm they sound like you.
- v0.28.0: an export from the desktop app and from the native window
  (TXT and PDF), replacing an existing file through the save dialog; and a
  look at the ten golden PDFs in `tests/golden/`.
- v0.29.1 on Windows and Linux: install and start on a real system,
  including the first-start warning (Windows) and the Linux menu entry.
  The Mac (install and **Update and restart**) is confirmed.
- The three decisions of 9 October marked "giopas to confirm" below.
- v0.30.0: a look at `tests/golden/marks-and-techniques.pdf` and
  `tests/golden/parts.pdf`; writing a chart with techniques, dynamics,
  a part and a riff in the app; exporting a PDF and opening it again with
  **Import song…**.

## Release history

Every version is in `CHANGELOG.md`, with the highlights table in
`ROADMAP.md`. This plan does not repeat them.

## Goal

A typed version of the one-page stage chart written by hand for practice:
fret numbers over chord symbols, repeat counts, named riffs, back
references and free annotations. `DESIGN_v0_16.md` section 1 holds the
vocabulary and the reference case.

Benchmark: a real song rebuilt with the tool only, printed on one page,
and readable on stage without the handwritten sheet.

## Principles

- The printed page is the reference. There is no other program to check
  against, so a change is judged by the PDF it prints next to the
  handwritten chart.
- Python standard library only. `pytest` for tests and `pywebview` for the
  native window are optional and never needed to run the app.
- Pure modules first, user interface last. Logic without Tkinter is tested
  without a display.
- Deterministic output: the same song (and the same footer date) gives the
  same TXT and PDF bytes.
- Nothing leaves the computer, apart from the update check of the installed
  app, which sends only its name and version and can be switched off.
  Lyrics are never fetched automatically.

## Standing decisions

| Date | Decision |
|---|---|
| Sep 2026 | The version lives only in `APP_VERSION` in `constants.py`. |
| Sep 2026 | `wiki/` in this repository is the wiki. A workflow publishes it when `main` changes; the wiki is never edited by hand. |
| Sep 2026 | Lyrics come from a local file or a paste. No automatic fetching from the web (copyright). |
| 8 Oct 2026 | Work follows the working agreements in `DEVELOPMENT.md` (branches, Conventional Commits, tests with every change, plain writing). |
| 8 Oct 2026 | Release notes are the version's `CHANGELOG.md` section. No separate release-notes files. |
| 8 Oct 2026 | The PDF footer keeps the print date. `build_pdf` takes an optional date so tests compare bytes. |
| 8 Oct 2026 | The `.sng` song file is saved in place, atomically, and renamed on Save without overwriting another song. Exports and imports never overwrite an existing file. |
| 8 Oct 2026 | The interface is English only for now. No translation checks in releases. |
| 8 Oct 2026 | Standalone packages for macOS, Windows and Linux are planned (item 8). Running from sources stays supported. |
| 8 Oct 2026 | App icon: option C, a fret number over a chord between a bar line and a repeat sign. Source `web/logo/icon.svg`. |
| 8 Oct 2026 | The package opens the native window (`webserver.py` with pywebview). The Tkinter desktop app stays available from sources. |
| 8 Oct 2026 | Packages get an update check with *Update and restart*, as in Swiss Knife. It is optional and disclosed, and sends nothing about the user. |
| 8 Oct 2026 | The apps are not signed or notarised. The README and wiki explain the first-run warning. |
| 8 Oct 2026 | Linux gets a `.tar.gz` with a `.desktop` launcher. No AppImage unless someone asks. |
| 9 Oct 2026 | The Linux package opens in the browser, because the native window needs GTK or Qt WebKit from the system. From the sources with pywebview, Linux gets the window. (Proposed by Claude while building item 8; giopas to confirm.) |
| 9 Oct 2026 | The update check runs by itself only in the installed app. From the sources, **Check now** in the About card asks on request. (Proposed by Claude while building item 8; giopas to confirm.) |
| 9 Oct 2026 | The release workflow leaves a draft release; giopas checks it and publishes it. Only published releases are offered by **Update and restart**. (Proposed by Claude while building item 8; giopas to confirm.) |
| 10 Oct 2026 | Techniques use the standard tab letters; palm mute is a run, `[pm …]`. Dynamics are lower-case words, an accent is `^`. (giopas, 10 October.) |
| 10 Oct 2026 | Parts are extra chart lines inside a section, each with its own instrument; ⚙ Layout picks which instruments print. (giopas, 10 October.) |
| 10 Oct 2026 | The song travels inside its PDF as an attachment, so a PDF can be opened again. TXT exports carry nothing. |
| 8 Oct 2026 | Claude has no role inside the app. It builds and checks the code only. |
| 8 Oct 2026 | Test data uses invented band, song and lyric names. The author's own songs stay out of the repository. |

## How we work

See `DEVELOPMENT.md`.

## Ideas not built

See the "Planned" and "Other ideas" lists in `ROADMAP.md`. An idea moves
into "What is left" here only when giopas picks it.
