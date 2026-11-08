# Work plan

The living plan for Song Notation Tool: where we are, what is left, and the
decisions in force. A new session reads this first, then `DEVELOPMENT.md`
for how the work is done. After a change here, the same file is copied to
the Claude project as `claude/WORKPLAN.md`.

This file is what the last session wrote down. Before relying on it, run
`python3 tools/ground_truth.py` and correct this file where the two
disagree. Finished items are deleted from here, not ticked: `CHANGELOG.md`
and git hold the history.

## Where we are

- Released: v0.33.0 and v0.33.1, published by the workflow on 8 November
  2026 with all eight files each. **Update and restart** from v0.33.0 to
  v0.33.1 failed: GitHub's download address
  (github.com/…/releases/download/…) answered 504 to every request, while
  the API address for the same file worked in under a second (checked
  from giopas's network).
- Ready to release: v0.33.2 on `fix/v0.33.2`: the updater downloads
  through the API address first, the download address as fallback. Not
  yet merged, tagged or pushed. Installed copies up to v0.33.1 cannot
  update themselves while GitHub's download address fails; install
  v0.33.2 by hand once.
- QLC+ Swiss Knife's `core/update.py` downloads the same way the old code
  here did (download address only, five seconds for SHA256SUMS), so it
  can fail the same way. Worth the same fix there.
- v0.32.0 was tagged and built on 8 November, but GitHub had retired the
  `macos-14` runner, so its draft has no `macos-arm64` file. Delete that
  draft instead of publishing it; v0.33.0 carries everything in it.
- Published before: v0.31.2 (13 October 2026).
- `release.yml` now builds Apple silicon on `macos-15` and publishes the
  release (no draft) only when every build has succeeded.
- Previous: v0.30.0, published on 10 October 2026 with all eight files
  (macOS Apple silicon and Intel, `.dmg` and `.tar.gz` each; Windows
  installer and `.zip`; Linux `.tar.gz`; `SHA256SUMS`).
- Every release from v0.26.3 on has its GitHub release.
- **Update and restart** is confirmed on the Mac (v0.29.0 to v0.29.1, on
  10 October).
- Test suite: 542 tests, all passing (`python3 -m pytest -q`).

## What is left

Housekeeping, small:

1. The three rewritten issue templates (from v0.27.1) are still not in
   `.github/ISSUE_TEMPLATE/`: commit `14fd35d` names them but only changed
   `release.yml`, and the files there are the originals from v0.15. Copy
   them in and commit them.
2. `feat/v0.30.0` was pushed to GitHub, against the working agreement.
   It is merged, so it can go: `git push origin --delete feat/v0.30.0`.

Still open from v0.31.0: a per-section size control in the desktop app,
if wanted, and giopas's check of a printed floor sheet.

The next piece of work after that is giopas's choice; see "Candidates for
the next version" below.

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

- v0.33.2: after installing it by hand, **Update and restart** to the
  next release on the Mac.
- v0.33.0: in the native window, add a Bass part to a guitar section, try
  each mode (Chart, Tab, Both, Free), export a PDF and look at the part;
  open an older song whose sections had a repeat and click the ×2 chip.

- v0.32.0: in the native window on the Mac, drag a section from the end
  of a long song to the top and back, press Escape halfway through a
  drag, move a section with the arrow keys on its handle, and copy one
  with ⧉ to the top and to the middle of the list.

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
- v0.30.0: **Update and restart** from v0.29.1 on the Mac.
- v0.31.0 floor sheet: print a real song with Size: Floor, lay it
  on the floor and read it from standing; say whether half-size tab is
  readable enough and whether the type is big enough. Try **Larger** and
  **Smaller** on a section, and **Recalled licks: Name only** on a song
  with a riff inside a referenced section.

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
| 13 Oct 2026 | Every GitHub release is titled `vX.Y.Z (a few words on the main change)`, for example `v0.31.2 (Floor sheet on one page)`. Since 8 November the few words go after the date in the version's CHANGELOG.md heading and the workflow sets the title. (giopas, 13 October.) |
| 8 Nov 2026 | A pushed tag publishes its release directly, as in Swiss Knife: no draft. `tools/release_notes.py` writes the title and text from CHANGELOG.md; the workflow publishes only when every build passed. (giopas, 8 November.) |
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
| 10 Oct 2026 | Techniques use the standard tab letters; palm mute is a run, `[pm …]`. Dynamics are lower-case words, an accent is `^`. (giopas, 10 October.) |
| 10 Oct 2026 | Parts are extra chart lines inside a section, each with its own instrument; ⚙ Layout picks which instruments print. (giopas, 10 October.) |
| 10 Oct 2026 | The song travels inside its PDF as an attachment, so a PDF can be opened again. TXT exports carry nothing. |
| 11 Oct 2026 | A session starts with `tools/ground_truth.py` and corrects this plan where it disagrees. Finished items are deleted from the plan, not ticked. Ideas borrowed from truongfelix/workflow; the rest of that template is not adopted. (giopas, 11 October.) |
| 13 Oct 2026 | A floor sheet: Size "Floor" prints the largest type that fits on one page (changed from two the same day), portrait or landscape, whichever is bigger; tab at about half size; a slim title (title, artist, key) and small section names. Each section can also be Smaller or Larger than the song. (giopas, 13 October.) |
| 13 Oct 2026 | Floor may use two columns when a long song gets clearly bigger type that way; a short song keeps one column across the full width. Tab never below 75%. (Proposed by Claude while building it; giopas to confirm.) |
| 8 Nov 2026 | Sections move by dragging their ⠿ handle (arrow keys on the handle from the keyboard); the ↑ ↓ buttons are gone. ⧉ asks where the copy goes and the copy stays a reference. Browser and native window only; the Tkinter app keeps its own drag and Ctrl/Cmd+D. (giopas, 8 November.) |
| 8 Nov 2026 | The README screenshots are retaken before every release with `tools/shots.py` (borrowed from Swiss Knife), and the pictures are committed with the release. (giopas, 8 November.) |
| 8 Nov 2026 | A part has the same four render modes as a section (Chart, Tab, Both, Free). The section repeat box is removed from the browser app; repeats go in the chart line, and an older whole-section repeat shows as a removable ×2 chip. (giopas, 8 November.) |
| 8 Oct 2026 | Claude has no role inside the app. It builds and checks the code only. |
| 8 Oct 2026 | Test data uses invented band, song and lyric names. The author's own songs stay out of the repository. |

## How we work

See `DEVELOPMENT.md`.

## Ideas not built

See the "Planned" and "Other ideas" lists in `ROADMAP.md`. An idea moves
into "What is left" here only when giopas picks it.
