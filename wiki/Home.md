# Song Notation Tool — Wiki

A typed version of the one-page stage chart the author writes by hand
for band practice: fret numbers over chord symbols, repeat counts,
named riffs, and back-references so a repeated chorus isn't written out
twice. Runs as a **desktop app**, a **browser front end**, or a
**headless CLI** — same engine, three surfaces.

## Pages

- **[[CLI and Web Use]]** — headless `cli.py` (convert / batch / lint)
  and the `webserver.py` browser front end: full command and API
  reference, beyond what's in the README.
- **[[Chart Line Syntax]]** — the compact grammar typed into a section's
  chart line: symbols, frets, repeats, groups, riffs, references, marks,
  and the free-text escape hatch.
- **[[Printing and Layout]]** — how a chart comes out on paper: section
  layout, colour vs black & white, fit-to-page scaling, printing through
  the OS, and where your songs and exports are kept.
- **[[Architecture]]** — how the pure engine (`model` / `grammar` /
  `render` / `transpose` / `songmap` / `export`) relates to the three
  front ends (`song_writer.py`, `cli.py`, `webserver.py`), and where to
  make a change once so all three pick it up.

## Elsewhere in the repo

These live as files in the repo itself, not the wiki, since they're
read by the app's own docstrings and by contributors working in an
editor rather than a browser:

- [README.md](https://github.com/giopas/song-notation-tool/blob/main/README.md)
  — quick start, features, file formats
- [CHANGELOG.md](https://github.com/giopas/song-notation-tool/blob/main/CHANGELOG.md)
  — what shipped in each version
- [ROADMAP.md](https://github.com/giopas/song-notation-tool/blob/main/ROADMAP.md)
  — what's planned next
- [DESIGN_v0_16.md](https://github.com/giopas/song-notation-tool/blob/main/DESIGN_v0_16.md)
  and [DESIGN_v0_17.md](https://github.com/giopas/song-notation-tool/blob/main/DESIGN_v0_17.md)
  — the data model, chart grammar, and song-map/riff-library design docs
  each release was built against
- [CONTRIBUTING.md](https://github.com/giopas/song-notation-tool/blob/main/CONTRIBUTING.md)
  — how to report bugs or suggest features
