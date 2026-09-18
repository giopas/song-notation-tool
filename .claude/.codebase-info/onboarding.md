# Onboarding

*Last Updated: 2026-09-18*

## Quick Start

```bash
# Desktop GUI (requires Tkinter)
python song_writer.py

# Web UI
python webserver.py          # opens http://localhost:8000

# CLI export
python cli.py convert songs/mysong.sng -f pdf
python cli.py batch songs/ -f txt
python cli.py lint songs/mysong.sng
```

No `pip install` needed for core functionality — everything runs on the Python 3.8+ stdlib.

For the native-window web mode, install pywebview:
```bash
pip install pywebview
```

## Running Tests

```bash
pytest                       # or: python -m pytest
```

All tests are in `tests/` and cover the six engine modules. No test fixtures require network
or GUI access.

## Common Tasks

### Add a new section type

1. Add the type name to `SECTION_TYPES` in `constants.py`
2. The Tkinter UI and web UI both read from this list dynamically

### Add a new instrument/tuning

1. Add to `INSTRUMENT_STRINGS` in `constants.py` (key = display name,
   value = number of strings)
2. The tab grid and export logic use this dict

### Add a new mark (barline, repeat sign, etc.)

1. Add the mark name to `MARK_NAMES` in `model.py`
2. Add lexer recognition in `grammar._lex()` and parsing in `grammar._parse_mark()`
3. Add rendering in `render.py` (how the mark appears in chart rows)
4. Add unparse logic in `grammar.unparse()` if needed
5. Add a test in `tests/test_grammar.py`

### Add a new item kind

1. Add to `ITEM_KINDS` in `model.py`
2. Add a `make_<kind>()` constructor in `model.py`
3. Add parser support in `grammar.py`
4. Add rendering in `render.py`
5. Update `songmap.py` if the new kind participates in references or usage tracking
6. Add tests

### Modify the document schema

1. Bump `FORMAT_VERSION` in `model.py`
2. Add a migration function `_migrate_vN_to_vN+1()` in `model.py`
3. Chain it in `migrate_document()`
4. Add a fixture and test in `tests/test_model.py`

### Export a song programmatically

```python
import json, model, export

with open("songs/mysong.sng") as f:
    doc = model.migrate_document(json.load(f))

txt = export.build_song_lines(doc)
pdf_bytes = export.build_pdf(doc, orient="portrait")
```

## File Format

`.sng` files are plain JSON. Key structure:
```json
{
  "format_version": 2,
  "meta": {"title": "...", "artist": "...", "key": "...", "time": "4/4", "bpm": "120"},
  "transpose": 0,
  "blocks": {"riff1": {"name": "Main Riff", "items": [...]}},
  "sections": [
    {
      "id": "verse1", "name": "Verse 1", "type": "Verse",
      "instrument": "Bass (4-string)", "render": "chart",
      "transpose": 0, "items": [...]
    }
  ]
}
```

## GitHub

Repository: `giopas/song-notation-tool`
