# Patterns & Conventions

*Last Updated: 2026-09-18*

## Design Patterns

### Pure Engine / Thin Frontend

Every computation — parsing, rendering, transposition, export, riff tracking, section
reordering — lives in a pure-Python module with no GUI imports. Frontends (`song_writer.py`,
`webserver.py`, `cli.py`) are thin shells that wire user actions to engine calls. This makes
the engine fully unit-testable without mocking any UI.

### Document-as-Dict

The document is a plain Python dict (JSON-serialisable) with a fixed schema defined in
`model.py`. There is no ORM or class hierarchy — items are dicts with a `kind` key, validated
by constructor functions (`make_token`, `make_group`, etc.) rather than class instantiation.

### Non-Destructive Transforms

Transposition is never written back into the document. The stored items keep their original
note/fret values; `transpose.resolve()` is called at render time with the effective semitone
offset. This keeps the source of truth stable and avoids rounding/alias drift.

### Item-Kind Dispatch

Both `grammar.py` (parsing) and `render.py` (rendering) dispatch on `item["kind"]`. The six
kinds — `token`, `group`, `block_ref`, `section_ref`, `measure`, `mark` — are enumerated in
`model.ITEM_KINDS`.

### Separate Chart and Measure Namespaces

A single section's `items` list can contain both chart-grammar items and tab-grid `measure`
items. `songmap.chart_items()` / `measure_items()` filter them apart so that committing an
edit in the chart editor never clobbers tab-grid data and vice versa
(`set_chart_items` / `set_measure_items`).

## Error Handling

- `grammar.ParseError` for malformed chart-line input (with position info)
- `ValueError` / `IndexError` for invalid operations in `songmap.py` (empty selection,
  out-of-range index)
- `model.validate_block_name()` raises `ValueError` for invalid riff identifiers
- The web server catches exceptions in handlers and returns JSON `{"error": message}` with
  appropriate HTTP status codes

## Testing

Tests live in `tests/` and mirror the engine modules 1:1:

| Test file            | Covers         |
|----------------------|----------------|
| `test_grammar.py`    | `grammar.py`   |
| `test_model.py`      | `model.py`     |
| `test_render.py`     | `render.py`    |
| `test_transpose.py`  | `transpose.py` |
| `test_songmap.py`    | `songmap.py`   |
| `test_export.py`     | `export.py`    |

Run with: `pytest` (or `python -m pytest` from the project root).

Test fixtures are in `tests/fixtures/` — notably `v015/sample_song.json` for migration tests.
Tests use plain `assert` statements (no unittest classes). Round-trip tests (parse → unparse →
parse) are a key pattern in `test_grammar.py`.

## Configuration

There is no config file or environment variable system. All tunables are constants:

- `constants.py` — app version, instrument definitions, section types
- `transpose.py` — chromatic scale, enharmonic map, MAX_FRET
- `render.py` — LINES_PER_PAGE, column widths
- `export.py` — page dimensions, font sizes, margins

## Coding Style

- `from __future__ import annotations` in every module
- Snake_case throughout; no classes in the engine except `namedtuple` (`Resolved`)
- Docstrings on public functions (triple-quote, imperative mood)
- No type annotations on function signatures (aside from `-> str` / `-> int` on small helpers)
- Imports are direct (`import model`, `import grammar`) — no package structure, all modules at
  top level
- No linter/formatter config files in the repo; style is consistent by convention
