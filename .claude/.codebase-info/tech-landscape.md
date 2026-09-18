# Tech Landscape

*Last Updated: 2026-09-18*

## Language & Runtime

| Aspect       | Value                                           |
|--------------|-------------------------------------------------|
| Language     | Python 3.8+ (uses `__future__.annotations`)     |
| GUI toolkit  | Tkinter (stdlib, desktop frontend only)          |
| Web server   | `http.server` (stdlib `ThreadingHTTPServer`)     |
| Optional     | `pywebview` — native window wrapper for web mode |
| Test runner  | `pytest`                                         |

## External Dependencies

**Core: zero.** The engine, CLI, and web server run on the Python standard library alone.

Optional runtime:
- `pywebview` — wraps the browser UI in a native OS window (resizable, titled).
  On macOS this pulls in PyObjC.

Dev only:
- `pytest` — test runner for the `tests/` suite.

## Source-of-Truth Files

| File              | What it governs                                           |
|-------------------|-----------------------------------------------------------|
| `constants.py`    | APP_VERSION, SECTION_TYPES, INSTRUMENT_STRINGS, defaults  |
| `model.py`        | FORMAT_VERSION, document schema, ITEM_KINDS, MARK_NAMES   |
| `grammar.py`      | Chart-line DSL lexer/parser/unparser                      |
| `transpose.py`    | CHROMATIC scale, ENHARMONIC aliases — single source of truth for note names |
| `.gitignore`      | Repo hygiene (no `.claude/*` exclusion)                   |
| `ROADMAP.md`      | Version history and planned features                      |
| `README.md`       | User-facing docs: install, run modes, CLI reference       |

## Version

Current release: **v0.19.0** (`constants.APP_VERSION`).
Document format: **version 2** (`model.FORMAT_VERSION`), with automatic migration from v1.
