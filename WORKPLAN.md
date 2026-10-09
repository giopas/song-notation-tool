# Work plan

The living plan for Song Notation Tool: where we are, what is left, and the
decisions in force. A new session reads this first, then `DEVELOPMENT.md`
for how the work is done. After a change here, the same file is copied to
the Claude project as `claude/WORKPLAN.md`.

## Where we are

- Released: v0.28.0 (one export writer, golden files), pushed on 9 October
  2026.
- Ready on `feat/v0.29.0` (9 October): item 8, the standalone packages and
  the update check. The Linux package was built and smoke-tested in
  Claude's workspace; macOS and Windows are built for the first time by the
  release workflow. Waiting for giopas to copy `release.yml` into
  `.github/workflows/`, then merge, tag and push.
- Test suite: 420 tests and 1 skipped (`python3 -m pytest -q`); the skipped
  one checks `release.yml` and runs once the file is in place.

## What is left

Work items, in order:

1. GitHub releases for v0.26.3, v0.27.0 and v0.27.1, each with its
   changelog section as the body (if not done yet).
2. Plain-language pass: done in v0.27.1, except the three issue
   templates in `.github/ISSUE_TEMPLATE/`. Claude cannot write under
   `.github/` on this computer (the folder is protected for remote
   tools), so the rewritten templates were handed to giopas to copy in.
   Left as they were on purpose: the ⚙ Layout option labels in
   `constants.py` ("All together — at the start" and so on), which are
   menu labels, and comments and docstrings in the code.
3. One writer: done in v0.28.0. `export.export_to_file()` (and
   `prepare_export()` for downloads) is the only place that builds and
   writes TXT and PDF exports. TXT is now the same bytes everywhere: UTF-8,
   Unix line endings, one newline at the end.
4. Exports never overwrite: done in v0.28.0. The command line keeps an
   existing file and writes `song_v2.pdf` unless `--force` is given. The
   save dialogs (desktop, native window) already ask before replacing, so
   they replace only when you agree.
5. Validate before export: done in v0.28.0. `export.check_document()`, the
   same check as `cli.py lint`, runs before every export in every front
   end. Errors stop it and name the section; chord-shape notes are shown.
6. Golden-file tests: done in v0.28.0. Ten invented songs in
   `tests/golden/`, built by `tools/make_golden.py`. PDFs are compared
   with their page streams inflated (zlib builds differ). Claude looked at
   all ten PDFs before committing them; giopas should look once too.
7. App icon: done in v0.27.0 (option C). Live check of how it looks in
   the Dock, the taskbar and a browser tab is on the manual list below.
8. Standalone packages: done in v0.29.0, as planned (spec, archives,
   Windows installer, Linux menu entry, release workflow with smoke tests
   and `SHA256SUMS`, update check with **Update and restart**, install
   sections in the README and the wiki). Still open: the first run of the
   release workflow on macOS and Windows (fix what it finds as v0.29.x),
   and the live checks below.

Manual checks waiting for giopas (each stays here until reported back):

- v0.26.3 repeat bracket in a printed PDF and in the web preview.
- Desktop app (Tkinter) and native window (pywebview) on the Mac after
  v0.27.0.
- Windows and Linux: never tried live. Built and unit-tested only.
- v0.27.0: the icon in the desktop app's window and Dock, in the native
  window, and as the favicon in a browser tab.
- v0.27.1: a read-through of the README and the wiki, and a look at
  the app's help texts, to confirm they sound like you.
- v0.28.0: an export from the desktop app and from the native window
  (TXT and PDF), replacing an existing file through the save dialog; and
  a look at the ten golden PDFs in `tests/golden/`.
- v0.29.0, after the release workflow has run: each package installed and
  started on a real system (macOS Apple silicon, macOS Intel, Windows,
  Linux), including the first-start warning of an unsigned app, and the
  Linux menu entry.
- v0.29.x: **Update and restart** from v0.29.0 to the next release, on the
  Mac first (Swiss Knife needed two fixes before it worked there).

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
| 8 Oct 2026 | Claude has no role inside the app. It builds and checks the code only. |
| 8 Oct 2026 | Test data uses invented band, song and lyric names. The author's own songs stay out of the repository. |

## How we work

See `DEVELOPMENT.md`.

## Ideas not built

See the "Planned" and "Other ideas" lists in `ROADMAP.md`. An idea moves
into "What is left" here only when giopas picks it.
