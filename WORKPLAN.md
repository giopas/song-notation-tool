# Work plan

The living plan for Song Notation Tool: where we are, what is left, and the
decisions in force. A new session reads this first, then `DEVELOPMENT.md`
for how the work is done. After a change here, the same file is copied to
the Claude project as `claude/WORKPLAN.md`.

## Where we are

- Released: v0.27.0 (the app icon), pushed on 8 October 2026 together
  with the v0.26.3 tag. The housekeeping planned as v0.26.4 shipped in
  v0.27.0.
- Test suite: 367 tests, all passing (`python3 -m pytest -q`).

## What is left

Work items, in order:

1. GitHub releases for v0.26.3 and v0.27.0, each with its changelog
   section as the body (if not done yet).
2. Plain-language pass (`feat/v0.27.1`), done before the packages so the
   install sections written for them start in the same voice. A complete
   review and rewrite with the `humanizer` skill of:
   - `README.md`, top to bottom (about 530 lines);
   - every page in `wiki/`;
   - the text the app shows, where it needs it: the help strip, "Start
     here" panel, tooltips and dialogs in `song_writer.py`, the `title=`
     hints, help and notation reference in `web/`, and the `cli.py` and
     `webserver.py` command-line help;
   - `CONTRIBUTING.md` and the issue templates.

   Rules for the pass: facts do not change, only the wording; anything
   that looks wrong is checked against the code and the changelog, and
   fixed there or reported, never invented. Button and menu names in the
   docs must match the app after the pass. Past `CHANGELOG.md` entries
   and the `DESIGN_v0_...md` files are history and stay as written. The
   writing style in `DEVELOPMENT.md` is the target. Tests that look for a
   user-visible string are updated with it.
3. One writer for files. TXT and PDF bytes come from `export.py`, but the
   files are written in three places (`cli.py`, `webserver.py`,
   `song_writer.py`). Move the writing into one function.
4. Exports never overwrite. The desktop app and the native window go
   through a save dialog, where the system asks before replacing a file.
   `cli.py convert` and `batch` write straight to the output path. Make
   them pick the next free name (`name_v2.pdf`) unless `--force` is given.
5. Validate before export. `cli.py lint` already checks that every chart
   line re-parses and warns about chord shapes. Run the same check before
   every export: errors block it, warnings are shown.
6. Golden-file tests: one anonymised song per layout feature, compared
   byte for byte (now possible since `build_pdf` takes a date).
7. App icon: done in v0.27.0 (option C). Live check of how it looks in
   the Dock, the taskbar and a browser tab is on the manual list below.
8. Standalone packages (`feat/v0.28.0`), built the way Swiss Knife 2.8
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
- After item 2: a read-through of the README and the wiki, and a look at
  the app's help texts, to confirm they sound like you.
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
