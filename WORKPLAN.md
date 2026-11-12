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

- Released: v0.34.0 (12 November 2026), tagged by giopas from `main`;
  the workflow builds and publishes it. Before it: v0.33.3 (8 November).
  Installed copies up to v0.33.1 cannot update themselves while GitHub's
  download address fails; install v0.33.2 or later by hand once.
- In progress: `feat/v0.35.0` (12 November), committed locally, not
  merged. Songs can be deleted (to a trash, restorable, **Empty trash…**),
  hidden, and put in collections (setlists, albums) with their own order
  and one PDF per collection. Browser and native window only; the Tkinter
  desktop app has no song list and is unchanged. `library.py` holds it;
  CHANGELOG `[Unreleased]`, README and wiki describe it. Before release:
  bump the version, screenshots (the song menu, a collection and the
  trash are worth a picture each in `tools/shots.py`), the pitch check.
- `chore/pitch-ci` and `fix/pitch-skip` are on `main` and pushed
  (ground truth, 12 November); both branches can be deleted. GitHub makes
  the pitch video and the deck for every release (`pitch.yml`, called by
  `release.yml`) and `pitch-page.yml` publishes the page that plays it
  (`https://giopas.github.io/song-notation-tool/`).
- 12 November: `pitch.yml` skipped a release's own run (called by
  `release.yml`, `GITHUB_EVENT_NAME` is the tag's `push`), so a release
  would never get its video; Barline 0.21.0 showed it. Fixed on
  `fix/pitch-skip`: it skips only on `main` with no tag given, and
  `pitch.yml` is in its own `push` paths, so pushing the fix makes the
  video for the latest release. Swiss Knife's `pitch.yml` needs the same
  check looked at, from that project.
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
- Test suite: 599 tests, all passing (`python3 -m pytest -q`).

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

- v0.35.0 (feat/v0.35.0) in the native window on the Mac: make a
  collection, add three songs from their ⋯ menu, drag one to the top and
  move one with Alt+↑; **Export as one PDF** and look at it; hide a song
  and find it under Hidden songs; delete the open song, Escape the first
  question, then delete it, restore it from Deleted songs and check it is
  back in its collection; **Empty trash…**. Then **Change…** the songs
  folder with a collection and a deleted song, and check both came along.
  Drag in the song list in the native window is untested (headless
  Chromium only).

- Pitch video: after pushing `chore/pitch-ci`, a green `pitch` run on the
  Actions page, `Song_Notation_Tool_Pitch.mp4` and `.pdf` on the v0.34.0
  release, a green `pitch-page` run, the README's poster opening
  `giopas.github.io/song-notation-tool` with the video playing, and a look
  at that page on a phone. A red run: **Run workflow** on the `pitch` page retries it.

- v0.34.0 in the native window on the Mac: open a song, check the new
  look in dark and light (the ◐ button), press Escape in each dialog,
  type a broken line (`[G D Em C x2`) and read the message, switch a
  section to Tab, add two measures and use +2, +4 and − on the second,
  open the Preview and check Save and Export stay visible, and make the
  window narrow to see the ☰ song list. Also open it on the phone over
  the home network (`--host 0.0.0.0`).
- v0.33.3: **Update and restart** from v0.33.2 on the Mac, then type
  `[5A 3C 5A 3G]x2 5A` in a section and check the frets sit over the
  chords in the preview and the PDF.
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
| 12 Nov 2026 | Deleting a song moves it to `Deleted songs` in the songs folder, restorable until **Empty trash…**; every delete asks first. Hidden songs and collections (plain, ordered, a song in several) live in `library.json` beside the songs, never in the song files. A collection exports as one PDF. (giopas, 12 November.) |
| 13 Oct 2026 | A floor sheet: Size "Floor" prints the largest type that fits on one page (changed from two the same day), portrait or landscape, whichever is bigger; tab at about half size; a slim title (title, artist, key) and small section names. Each section can also be Smaller or Larger than the song. (giopas, 13 October.) |
| 13 Oct 2026 | Floor may use two columns when a long song gets clearly bigger type that way; a short song keeps one column across the full width. Tab never below 75%. (Proposed by Claude while building it; giopas to confirm.) |
| 8 Nov 2026 | Sections move by dragging their ⠿ handle (arrow keys on the handle from the keyboard); the ↑ ↓ buttons are gone. ⧉ asks where the copy goes and the copy stays a reference. Browser and native window only; the Tkinter app keeps its own drag and Ctrl/Cmd+D. (giopas, 8 November.) |
| 8 Nov 2026 | The README screenshots are retaken before every release with `tools/shots.py` (borrowed from Swiss Knife), and the pictures are committed with the release. (giopas, 8 November.) |
| 8 Nov 2026 | A part has the same four render modes as a section (Chart, Tab, Both, Free). The section repeat box is removed from the browser app; repeats go in the chart line, and an older whole-section repeat shows as a removable ×2 chip. (giopas, 8 November.) |
| 11 Nov 2026 | Every release includes a pitch check: does the pitch video and slide deck (`tools/make_pitch_video.py`) still tell the truth about the app? Claude reads the version's CHANGELOG.md section against `SLIDES` (a claim now wrong, a feature renamed or removed, a new feature worth a bullet or a slide, a slide's narration that no longer matches), against the `tools/shots.py` report (a retaken screenshot whose zoom point `name.png@x,y,zoom` now lands somewhere else), and against the golden PDFs (a change to the printed output means `--refresh-pages`). It then runs `--stills-only` and looks at the slides. If something needs to change, Claude proposes the exact edit to the script (or the refreshed pictures) with the release, for giopas to approve; if nothing does, the release summary says "Pitch video: no change". The video and deck are always made by giopas on the Mac with exactly `python3 make_pitch_video.py --stills --tts edge` (run from `tools/`): the edge-tts voice for every slide, plus the PNG stills and the PDF deck. Claude gives that command in every release summary whose pitch check changed something, and never proposes another voice or option. The output is never committed (`dist/` is ignored). (giopas, 11 November.) |
| 11 Nov 2026 | The pitch video's first slide shows a photo of one of giopas's handwritten charts (`tools/pitch_media/paper-handwritten.jpg`), as the "before" next to the app's PDF. This is the one exception to keeping the author's own songs out of the repository. (giopas, 11 November.) |
| 12 Nov 2026 | GitHub makes the pitch video and the slide deck and attaches them to every release (`pitch.yml`, after `release.yml` publishes), always with the edge voice (`--strict-voice`: no video without it). Nobody makes or uploads them by hand any more; `make_pitch_video.py --stills --tts edge` on the Mac stays for a look before a release. The README links to the latest release's files. The pitch check itself (Claude proposes, giopas approves the words and zoom points) is unchanged. This replaces the 11 November rule that giopas makes the video on the Mac. Adapted from QLC+ Swiss Knife. (giopas, 12 November.) |
| 8 Oct 2026 | Claude has no role inside the app. It builds and checks the code only. |
| 8 Oct 2026 | Test data uses invented band, song and lyric names. The author's own songs stay out of the repository. |

## How we work

See `DEVELOPMENT.md`.

## Ideas not built

See the "Planned" and "Other ideas" lists in `ROADMAP.md`. An idea moves
into "What is left" here only when giopas picks it.
