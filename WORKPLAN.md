# Work plan

The living plan for Song Notation Tool: where we are, what is left, and the
decisions in force. A new session reads this first, then `DEVELOPMENT.md`
for how the work is done. After a change here, the same file is copied to
the Claude project as `claude/WORKPLAN.md`.

## Where we are

- Released: v0.27.1 (the plain-language pass), pushed on 9 October 2026.
- Ready on `feat/v0.28.0` (9 October): items 3 to 6, one export writer
  that checks every song first and never overwrites, and golden-file
  tests. Waiting for merge, tag and push.
- Test suite: 398 tests, all passing (`python3 -m pytest -q`), including
  ten golden songs.

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
8. Standalone packages (`feat/v0.29.0`), built the way Swiss Knife 2.8
   does it:
   - a PyInstaller spec (`packaging/songnotation.spec`) and
     `packaging/make_archive.py`, which names the files
     `Song-Notation-Tool-<version>-<os>-<arch>.<ext>`;
   - a `release.yml` workflow that runs on a version tag and builds on
     macOS Apple silicon, macOS Intel, Windows and Linux, runs the tests,
     smoke-tests each bundle (`--smoke`), and attaches the files and
     `SHA256SUMS` to the GitHub release;
   - macOS: a `.app` in a `.dmg` (and a `.zip`); Windows: an Inno Setup
     installer and a `.zip`, with the installer tested by a silent
     install, start and uninstall in the workflow; Linux: a `.tar.gz`
     with a `.desktop` launcher;
   - README and wiki: an install section per system, including the
     unsigned-app warning and how to get past it;
   - an update check against the GitHub releases, with *Update and
     restart*, as in Swiss Knife (`core/update.py` there);
   - running from sources keeps working exactly as today.

   Note: `.github/` is protected for Claude's tools on this computer, so
   `release.yml` will be written in the session's own workspace and
   copied into `.github/workflows/` by giopas.

   Decided on 8 October (see standing decisions): the package opens the
   native window; it has an update check with *Update and restart*, as
   in Swiss Knife; no signing or notarisation; Linux gets `.tar.gz`
   only, no AppImage.

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
- After item 8: each package installed and started on a real system
  (macOS Apple silicon, macOS Intel, Windows, Linux).

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
- Nothing leaves the computer. Lyrics are never fetched automatically.

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
| 8 Oct 2026 | Claude has no role inside the app. It builds and checks the code only. |
| 8 Oct 2026 | Test data uses invented band, song and lyric names. The author's own songs stay out of the repository. |

## How we work

See `DEVELOPMENT.md`.

## Ideas not built

See the "Planned" and "Other ideas" lists in `ROADMAP.md`. An idea moves
into "What is left" here only when giopas picks it.
