# Entry Points

*Last Updated: 2026-09-18*

## Desktop GUI

| Entry               | How to run                  | What happens                             |
|----------------------|-----------------------------|------------------------------------------|
| `song_writer.py`     | `python song_writer.py`     | Launches the Tkinter desktop application |

The Tkinter app is the primary authoring surface. It owns the chart editor bar, tab grid,
song map, riff library panel, and all keyboard shortcuts.

## Web / Browser UI

| Entry               | How to run                        | What happens                                    |
|----------------------|-----------------------------------|-------------------------------------------------|
| `webserver.py`       | `python webserver.py`             | Starts stdlib HTTP server on `localhost:8000`    |
| with pywebview       | `python webserver.py`             | Same, but opens a native OS window via pywebview |

The server auto-opens `http://localhost:8000` in the default browser unless pywebview is
installed, in which case it opens a titled, resizable native window instead.

### REST API Routes (webserver.py)

| Method | Path                              | Purpose                        |
|--------|-----------------------------------|--------------------------------|
| GET    | `/api/meta`                       | App version + instruments list |
| GET    | `/api/songs`                      | List saved .sng files          |
| GET    | `/api/songs/<name>`               | Load a song                    |
| GET    | `/api/songs/<name>/export.txt`    | Download TXT export            |
| GET    | `/api/songs/<name>/export.pdf`    | Download PDF export            |
| POST   | `/api/songs`                      | Create new song                |
| POST   | `/api/songs/example`              | Create example song            |
| POST   | `/api/parse`                      | Parse a chart line             |
| POST   | `/api/render`                     | Render items to chart row      |
| POST   | `/api/export.txt`                 | Export current doc as TXT      |
| POST   | `/api/export.pdf`                 | Export current doc as PDF      |
| PUT    | `/api/songs/<name>`               | Update/save a song             |
| DELETE | `/api/songs/<name>`               | Delete a song file             |
| POST   | `/api/quit`                       | Graceful server shutdown       |

Static files under `web/` are served for any path not matching `/api/`.

## Headless CLI

| Entry     | How to run                                      | What happens                          |
|-----------|-------------------------------------------------|---------------------------------------|
| `cli.py`  | `python cli.py convert song.sng -f pdf`         | Convert one .sng to TXT or PDF        |
|           | `python cli.py batch ./songs/ -f txt`           | Batch-convert a folder of .sng files  |
|           | `python cli.py lint song.sng`                   | Round-trip parse check (lint)         |

All three subcommands share a common argparse parent for `--instruments` and `--orient` flags.
