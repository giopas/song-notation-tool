<img src="https://raw.githubusercontent.com/giopas/song-notation-tool/main/web/logo/icon-256.png" width="96" height="96" alt="">

# Song Notation Tool wiki

A typed version of the one-page stage chart the author writes by hand for
band practice: fret numbers over chord symbols, repeat counts, named
riffs, and back-references so a repeated chorus is not written out twice.
It runs as a desktop app, a browser front end or a command-line tool, all
on the same engine.

## Pages

- [[CLI and Web Use]]: the command-line `cli.py` (convert, batch, lint) and
  the `webserver.py` browser front end, with the full command and API
  reference.
- [[Chart Line Syntax]]: what you type on a section's chart line. Symbols,
  frets, repeats, groups, riffs, references, licks (and naming and
  recalling them), marks, and free text for what the grammar does not
  cover.
- [[Printing and Layout]]: how a chart comes out on paper, all set from
  **⚙ Layout** (size, columns, colour, section names, where the lyrics go,
  licks and chord shapes); printing through the system; and where your
  songs and exports are kept, including changing and moving the songs
  folder.
- [[Architecture]]: how the engine (`model`, `grammar`, `render`,
  `transpose`, `songmap`, `export`) relates to the three front ends
  (`song_writer.py`, `cli.py`, `webserver.py`), and where to make a change
  once so all three get it.

## Elsewhere in the repository

These are files in the repository, read by contributors working in an
editor:

- [README.md](https://github.com/giopas/song-notation-tool/blob/main/README.md):
  quick start, features, file formats
- [CHANGELOG.md](https://github.com/giopas/song-notation-tool/blob/main/CHANGELOG.md):
  what changed in each version
- [ROADMAP.md](https://github.com/giopas/song-notation-tool/blob/main/ROADMAP.md):
  what is planned
- [DESIGN_v0_16.md](https://github.com/giopas/song-notation-tool/blob/main/DESIGN_v0_16.md)
  and [DESIGN_v0_17.md](https://github.com/giopas/song-notation-tool/blob/main/DESIGN_v0_17.md):
  the design documents for the data model, the chart grammar, the song map
  and the riff library
- [CONTRIBUTING.md](https://github.com/giopas/song-notation-tool/blob/main/CONTRIBUTING.md):
  how to report a bug or suggest a feature

These pages are published from `wiki/` in the
[main repository](https://github.com/giopas/song-notation-tool). An edit
made in the wiki's own editor is overwritten the next time it publishes,
so send a pull request against `wiki/` instead.
