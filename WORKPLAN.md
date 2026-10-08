# Work plan

The living plan for Song Notation Tool: where we are, what is left, and the
decisions in force. A new session reads this first, then `DEVELOPMENT.md`
for how the work is done. After a change here, the same file is copied to
the Claude project as `claude/WORKPLAN.md`.

## Where we are

- Released: v0.26.2 (tagged and pushed). v0.26.3 is on `main` on GitHub and
  in the changelog, but its tag was only created on 8 October 2026 and is
  not pushed yet. No GitHub release exists for it.
- In progress: `fix/v0.26.4`. Repository hygiene and a deterministic PDF,
  from adopting the working agreements (see the changelog `[Unreleased]`).
- Test suite: 362 tests, all passing (`python3 -m pytest -q`).

## What is left

Work items, in order:

1. Push the v0.26.3 tag and publish its GitHub release from the changelog.
2. Finish v0.26.4: merge, tag, push, release.
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

Manual checks waiting for giopas (each stays here until reported back):

- v0.26.3 repeat bracket in a printed PDF and in the web preview.
- Desktop app (Tkinter) and native window (pywebview) on the Mac after
  v0.26.4.
- Windows and Linux: never tried live. Built and unit-tested only.

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
| 8 Oct 2026 | Runs from source. No installer, packages or update check. |
| 8 Oct 2026 | Claude has no role inside the app. It builds and checks the code only. |
| 8 Oct 2026 | Test data uses invented band, song and lyric names. The author's own songs stay out of the repository. |

## How we work

See `DEVELOPMENT.md`.

## Ideas not built

See the "Planned" and "Other ideas" lists in `ROADMAP.md`. An idea moves
into "What is left" here only when giopas picks it.
