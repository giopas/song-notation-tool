# Directory Structure

*Last Updated: 2026-09-18*

```
song-notation-tool/
├── song_writer.py          # Tkinter desktop app (~2000 lines, main GUI)
├── cli.py                  # Headless CLI (convert / batch / lint)
├── webserver.py            # Stdlib HTTP server + optional pywebview window
├── model.py                # Document schema, item constructors, migration
├── grammar.py              # Chart-line DSL parser / unparser
├── render.py               # Shared rendering (column-aligned chart rows)
├── transpose.py            # Chromatic scale, enharmonic map, transposition
├── songmap.py              # Song-map helpers (riff usage, reorder, promote)
├── export.py               # TXT and hand-built PDF export engine
├── constants.py            # APP_VERSION, instruments, section types
├── examples.py             # Built-in "Example Song" document
├── web/
│   ├── index.html          # Single-page browser UI shell
│   ├── style.css           # Browser UI styles
│   └── app.js              # Browser front-end JS (~712 lines)
├── tests/
│   ├── test_export.py      # TXT/PDF export tests
│   ├── test_grammar.py     # Parser round-trip and edge-case tests
│   ├── test_model.py       # Schema, migration, validation tests
│   ├── test_render.py      # Chart-row rendering tests
│   ├── test_songmap.py     # Riff usage, reorder, promote tests
│   ├── test_transpose.py   # Transposition and enharmonic tests
│   └── fixtures/
│       └── v015/
│           └── sample_song.json  # v1-format fixture for migration tests
├── README.md               # User documentation
├── ROADMAP.md              # Version history and planned features
├── .gitignore
└── .claude/
    └── .codebase-info/     # This codebase map (you are here)
```

### Local-only directories (not in git)

| Directory                 | Contents                                      |
|---------------------------|-----------------------------------------------|
| `Old - do not commit/`    | Archived source from v0.1–v0.15               |
| `.venv/`                  | Python 3.14 virtualenv (pywebview, PyObjC)    |
| `songs/`                  | User's `.sng` data files                      |
| `song-notation-tool.wiki/`| Clone of the GitHub wiki                      |
| `wiki/`                   | Duplicate of wiki content                     |

### Organising Principle

**Layer-based by role**: top-level `.py` files are the engine + frontends, `web/` holds browser
assets, `tests/` holds pytest tests mirroring the engine modules 1:1.
