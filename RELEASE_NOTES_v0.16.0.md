# Song Notation Tool v0.16.0 — Foundation release

This is the first tagged release of Song Notation Tool, and the foundation
for everything the project is actually for: reproducing the compact,
one-page stage charts the author writes by hand.

## Why

v0.15 supported exactly one thing — a full measure-by-measure tab grid —
which meant every song came out as several pages, mostly dashes, with
repeated sections written out in full. See `DESIGN_v0_16.md` section 2 for
the full audit against a real song (`examples/sample_song.sng`).

## Highlights

- **New data model** (format 2): a section is a flat sequence of items —
  token, group, block reference, section reference, measure, mark — instead
  of four parallel per-measure layers.
- **Chart grammar**: type `5A 5D [5A 5D]x2 riff1 x3` instead of filling in a
  measure grid cell by cell.
- **Non-destructive transposition**: `document + section + reference`
  offsets, applied at render time. Transposing up and back down returns the
  file to its exact original state.
- **Old files still load.** A v0.15 `.sng` migrates automatically —
  `link_id` becomes a `section_ref`, so linked sections (like a repeated
  chorus) stop being duplicated in the output.
- **Bug fixes**: a duplicate-key bug in the enharmonic note table, TXT/PDF
  disagreeing on which sections have tab content, an unused text-wrap
  width that let every section export as one unreadable line, and fret
  entry silently accepting invalid values like `01`.

## Upgrading

Open any `.sng` file saved by v0.15 — it loads as-is. Nothing is
destructive; the app will not overwrite your file with the new format
until you save.

## What's not in this release

The song-map whole-song view, the riff-library panel, the chart-row inline
editor with live re-render, mark/ending rendering, one-page pagination, and
the stage/print view are deferred to v0.17 — see `ROADMAP.md`.

## Try it

```bash
python3 song_writer.py
```

`examples/sample_song.sng` is a real migrated file if you want to see the
new format without setting anything up from scratch.
