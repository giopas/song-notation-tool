# Architecture

*Last Updated: 2026-09-18*

## Overview

Song Notation Tool is a compact one-page song charting application for bass and guitar. It lets
musicians jot down how to play a song using a mix of chord names, fret numbers, tab grids,
repetition marks, and riff references — all rendered as dense, column-aligned text suitable for
a single printed page.

## Three-Frontend / Shared-Engine Design

```
┌─────────────────┐  ┌──────────────┐  ┌─────────────────────┐
│  Tkinter Desktop │  │  Headless CLI │  │  Web (browser + opt  │
│  song_writer.py  │  │  cli.py       │  │  pywebview window)   │
│  (~2000 lines)   │  │  (~186 lines) │  │  webserver.py + web/ │
└────────┬─────────┘  └──────┬────────┘  └──────────┬──────────┘
         │                   │                      │
         └───────────┬───────┴──────────────────────┘
                     │
        ┌────────────▼────────────────┐
        │     Pure-Python Engine      │
        │  model / grammar / render   │
        │  transpose / songmap        │
        │  export / constants         │
        │  examples                   │
        │  (zero external deps)       │
        └─────────────────────────────┘
```

All three frontends share the same engine modules. The engine is pure Python with **zero external
dependencies** (PDF export uses only `zlib`), making the core importable and testable without any
GUI toolkit.

## Data Flow

1. **Load** — A `.sng` file (plain JSON, format version 2) is read and optionally migrated from
   v1 (v0.15 schema) via `model.migrate_document()`.
2. **Edit** — The frontend manipulates the document dict in memory. Chart lines are strings in a
   custom DSL parsed by `grammar.py` into item lists (tokens, groups, marks, refs).
3. **Render** — `render.py` resolves display items (applying non-destructive transposition via
   `transpose.py`) and column-aligns fret/symbol pairs into two-line chart rows.
4. **Export** — `export.py` produces TXT (plain text) or PDF (hand-built byte stream with
   Helvetica/Courier, zlib-compressed) from the rendered lines.
5. **Save** — The document dict is serialised back to `.sng` JSON.

## Key Architectural Decisions

- **Non-destructive transposition**: effective semitones = document + section + ref transpose,
  computed at render time only (`transpose.effective_transpose`). The stored document is never
  mutated by transposition.
- **Riff library with usage tracking**: Named blocks (riffs) are stored in `doc["blocks"]` and
  referenced by `block_ref` items. `songmap.block_usage()` prevents deletion of referenced riffs.
- **Section references over copies**: "Duplicate as reference" (`songmap.duplicate_as_reference`)
  inserts a `section_ref` pointing at the original, avoiding the stale-copy problem.
- **Chart items vs. measure items**: A section's `items` list can hold both chart-grammar items
  and tab-grid `measure` items. `songmap.chart_items()` / `measure_items()` keep them apart so
  editing one surface never clobbers the other.
- **Hand-rolled PDF**: `export._assemble_pdf()` builds PDF objects directly (ported from QLC+
  Swiss Knife), avoiding any dependency on reportlab or similar.
