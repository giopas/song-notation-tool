# 🎸 Song Notation Tool

> **⚠️ DRAFT / DEMO — Vibecoded Prototype**
> This is an early-stage personal project built through iterative AI-assisted ("vibecoded") development. It is functional but rough around the edges. Use it, break it, and feel free to suggest improvements!

---

## What this produces

A one-page stage chart, in the same compact vocabulary the author already
uses on paper: fret numbers over chord symbols, repeat counts, named riffs
referenced by name, and back-references so a repeated chorus isn't written
out twice.


Built entirely with Python's standard library (Tkinter), so **no pip
installs required**. Compatible with **macOS**, **Windows**, and **Linux**.

---

## ✨ Features

- **Chart grammar** — type a line like `5A 5D [5A 5D]x2 riff1 x3` instead
  of filling in a measure grid cell by cell
- **Section-based layout** — Intro, Verse, Chorus, Bridge, Solo, and more,
  arranged in a scrollable left panel
- **Four notation layers per section** — Tab, Chords, Notes, Lyrics — each
  togglable independently
- **Guitar & Bass tunings** — 6-string, 7-string, 4/5-string bass, with
  Drop D variants
- **Variable beats per measure** — 8, 16, 32, or 64 beats; configurable per
  section
- **Non-destructive transposition** — shift the whole song or a section by
  any number of semitones; transposing up and back down is exact
- **Section linking / references** — a repeated section doesn't duplicate
  its content in the saved file or the export
- **Compact export** — empty strings are dropped, references aren't
  expanded, and PDF targets one page (Portrait A4 default)
- **Copy & paste sections** — duplicate any section with a single dialog
- **Light / Dark theme** — switch on the fly
- **Export** — save as `.txt` (plain text) or `.pdf` (formatted sheet with
  footer)
- **Project files** — save/load sessions as `.sng` (plain JSON,
  human-readable); old `.sng` files from earlier versions load without loss
- **Zero dependencies** — pure Python standard library, works out of the box

---

## 📸 Screenshots

![Song Notation Tool](screenshots/song_notation_screenshot.png)

---

## 🚀 Quick Start

```bash
# No install needed — just run:
python3 song_writer.py
```

Requires Python 3.8+ with Tkinter (included in most Python distributions).

> On some Linux distros Tkinter is a separate package:
> ```bash
> sudo apt install python3-tk
> ```

---

## 📁 File Formats

| Extension | Description |
|---|---|
| `.sng` | Project file — plain JSON, stores all sections and layers |
| `.txt` | Plain-text export, readable in any editor |
| `.pdf` | Formatted export with title, artist, date footer |


---

## 🗺️ Roadmap

See [ROADMAP.md](ROADMAP.md) for planned features, and
[CHANGELOG.md](CHANGELOG.md) for what's already shipped.

---

## 🛠️ Development

See [DEVELOPMENT.md](DEVELOPMENT.md) for how sessions on this project are
run, and the current `DESIGN_v0_16.md` for the data model, chart grammar,
and transposition rules this release is built on.

Run the test suite with:

```bash
pip install pytest
python3 -m pytest tests/
```

---

## 🤝 Contributing

This is a personal / demo project for now, but feedback and ideas are very welcome!
Check [CONTRIBUTING.md](CONTRIBUTING.md) for how to report bugs or suggest features.

---

## 📄 License

MIT License — see [LICENSE](LICENSE).

> **Note:** This project is an independent personal tool and is not affiliated with, endorsed by, or connected to any commercial product or organisation.
