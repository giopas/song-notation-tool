#!/usr/bin/env python3
"""
webserver.py — Song Notation Tool local web server.

A thin stdlib http.server layer over the same pure engine the Tkinter
app and the CLI use (model / grammar / render / transpose / songmap /
export) — no Tkinter import, and no *required* external dependencies
(the project's "zero external dependencies" rule extends here too).
This is the browser-front-end half of "Enhancement 2 — browser and
headless use" in the UX & Multi-Platform spec: same sections/grid
editing model as the desktop app, reachable at
http://<this-machine>:PORT from any device on the network (e.g. a
phone or tablet at rehearsal).

Run it:

    ./webserver.py                 # opens a window automatically —
                                    # a native app window (resizable,
                                    # not a browser tab) if pywebview
                                    # is installed, else your default
                                    # browser (with install
                                    # instructions printed)
    ./webserver.py --dir mysongs --port 9000
    ./webserver.py --browser       # force a normal browser tab, even
                                    # if pywebview is installed
    ./webserver.py --no-open       # just start the server, open
                                    # nothing (e.g. when only serving
                                    # other devices on the network)
    ./webserver.py --host 0.0.0.0  # reachable from other devices on
                                    # the LAN

pywebview (`pip install pywebview`, or `pip install -r
requirements-optional.txt`) is entirely optional — like the sibling
qlc-plus-swiss-knife-tool-script project, its native-window mode is
used automatically when it's importable and falls straight back to
opening your default browser tab when it isn't, printing install
instructions when that fallback is due to the missing package rather
than an explicit --browser. Nothing here requires it to run. The
window is a normal titled OS window — resizable, with its own
minimize/maximize/close buttons — not frameless: that was tried first
and reverted once testing showed a borderless pywebview window loses
native resize-by-edge and maximize on macOS, with nothing pywebview
exposes to reliably replace either. The front end's "Save & Close"
button is still there as a convenience alongside the window's own
close button — it saves the open song, then calls POST /api/quit,
which destroys the native window (or, in browser mode, stops the
HTTP server via socketserver's own thread-safe shutdown()) — same
idea as that sibling project's own /api/quit route.

Songs are plain .sng JSON files (the same format the desktop app saves)
read from/written to --dir. The API is documented inline below, next to
each route.
"""

from __future__ import annotations

import argparse
import json
import mimetypes
import os
import re
import socketserver
import subprocess
import sys
import threading
import webbrowser
from http import HTTPStatus
from http.server import BaseHTTPRequestHandler
from urllib.parse import urlparse, parse_qs, unquote

import chords as chords_mod
import export
import grammar
import userpaths
import lyrics as lyrics_mod
import model
import songmap
import transpose
import render as render_mod
from examples import example_document
from constants import (
    APP_VERSION, INSTRUMENT_STRINGS, SECTION_TYPES, RENDER_MODE_LABELS,
    SECTION_LAYOUT_LABELS, COLOR_MODE_LABELS, PDF_SCALE_LABELS,
    PDF_COLUMN_LABELS, LYRICS_LAYOUT_LABELS, CHORD_SHEET_LABELS,
    LICK_REF_LABELS,
    TAB_BEATS_DEFAULT, default_export_name,
)

WEB_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "web")

# Set by main() so /api/quit can shut down whichever mode is actually
# running: destroy the native window if pywebview is in use, otherwise
# stop the HTTP server directly via its own thread-safe shutdown() —
# same idea as the sibling qlc-plus-swiss-knife-tool-script project's
# /api/quit route, but via socketserver's built-in stop mechanism rather
# than self-signaling, which needs no OS signal-delivery step at all and
# so works the same in every environment.
_RUNTIME = {"webview_window": None, "httpd": None}


class _JsApi:
    """Exposed to the front end as `window.pywebview.api` in native-window
    mode. `window.open(url, "_blank")` is a no-op there (pywebview's
    webview has no browser-tab concept to open it in), so the Lyrics
    panel's "Search online" button calls this instead — same
    webbrowser.open_new_tab() the desktop app's own lyrics search uses
    — to hand the URL to the OS's actual default browser. In plain
    browser mode (--browser, or pywebview not installed) `window.pywebview`
    doesn't exist at all and the front end falls back to window.open()."""

    def open_url(self, url: str) -> bool:
        if not isinstance(url, str) or not url.startswith(("http://", "https://")):
            return False
        webbrowser.open_new_tab(url)
        return True

    # ── Export ────────────────────────────────────────────────────────
    # In a browser tab an export is a download and the browser decides
    # where it lands. In the native window there's no download UI at all,
    # so without this an export would either vanish or land somewhere the
    # user never chose. Here the document is rendered in Python and written
    # to a path the OS save panel returned — so "where did it go?" is
    # answered before the file is written, not after.
    def export_document(self, doc: dict, fmt: str = "txt", orient: str = "portrait"):
        import webview
        try:
            doc = model.migrate_document(doc or {})
            fmt = "pdf" if str(fmt).startswith("pdf") else "txt"
            default_name = default_export_name(doc, fmt)

            window = _RUNTIME.get("webview_window")
            if window is None:
                return {"ok": False, "error": "no window"}
            chosen = window.create_file_dialog(
                webview.SAVE_DIALOG,
                directory=userpaths.last_export_dir(),
                save_filename=default_name,
            )
            if not chosen:
                return {"ok": False, "cancelled": True}
            path = chosen if isinstance(chosen, str) else chosen[0]

            if fmt == "txt":
                data = ("\n".join(export.build_song_lines(doc)) + "\n").encode("utf-8")
            else:
                data = export.build_pdf(doc, orient=orient)
            with open(path, "wb") as f:
                f.write(data)

            userpaths.set_last_export_dir(os.path.dirname(path))
            return {"ok": True, "path": path}
        except Exception as exc:  # noqa: BLE001 — surfaced in the UI, never fatal
            return {"ok": False, "error": str(exc)}

    def print_document(self, doc: dict, orient: str = "portrait"):
        """
        Print through the operating system's own print dialog.

        Deliberately not a silent `lp` job: the PDF is written to a temp
        file and handed to the OS default viewer, so the user gets the
        real print panel — printer, paper size, scaling, page range — and
        a preview of what they're about to put on paper. A stage chart is
        exactly the kind of thing you want to eyeball before printing.
        """
        import tempfile
        try:
            doc = model.migrate_document(doc or {})
            data = export.build_pdf(doc, orient=orient)
            name = default_export_name(doc, "pdf")
            path = os.path.join(tempfile.mkdtemp(prefix="song-notation-"), name)
            with open(path, "wb") as f:
                f.write(data)

            if sys.platform == "darwin":
                subprocess.Popen(["open", path])
            elif os.name == "nt":
                try:
                    os.startfile(path, "print")  # noqa: S606
                except OSError:
                    os.startfile(path)           # noqa: S606
            else:
                subprocess.Popen(["xdg-open", path])
            return {"ok": True, "path": path}
        except Exception as exc:  # noqa: BLE001
            return {"ok": False, "error": str(exc)}

    # ── Songs folder ──────────────────────────────────────────────────
    def choose_songs_folder(self):
        """Pick a new songs folder. Existing songs are left where they
        are — this points the app at a folder, it doesn't move data."""
        import webview
        try:
            window = _RUNTIME.get("webview_window")
            if window is None:
                return {"ok": False, "error": "no window"}
            chosen = window.create_file_dialog(
                webview.FOLDER_DIALOG, directory=userpaths.songs_dir())
            if not chosen:
                return {"ok": False, "cancelled": True}
            path = chosen if isinstance(chosen, str) else chosen[0]
            userpaths.set_songs_dir(path)
            Handler.store = SongStore(path)
            return {"ok": True, "path": path}
        except Exception as exc:  # noqa: BLE001
            return {"ok": False, "error": str(exc)}

    def reveal(self, path: str = "") -> bool:
        """Open a folder in Finder / Explorer / the desktop file manager."""
        target = os.path.abspath(os.path.expanduser(path or userpaths.songs_dir()))
        if not os.path.isdir(target):
            target = os.path.dirname(target)
        try:
            if sys.platform == "darwin":
                subprocess.Popen(["open", target])
            elif os.name == "nt":
                os.startfile(target)  # noqa: S606
            else:
                subprocess.Popen(["xdg-open", target])
            return True
        except Exception:  # noqa: BLE001
            return False


def _shutdown():
    window = _RUNTIME.get("webview_window")
    if window is not None:
        try:
            window.destroy()
        except Exception:
            os._exit(0)
        return
    httpd = _RUNTIME.get("httpd")
    if httpd is not None:
        # shutdown() blocks until serve_forever() (running on the main
        # thread) notices and returns — call it from its own thread so
        # this timer callback doesn't block waiting for itself.
        threading.Thread(target=httpd.shutdown, daemon=True).start()
    else:
        os._exit(0)

_SAFE_NAME_RE = re.compile(r"[^A-Za-z0-9 _.-]")


def _safe_filename(name: str) -> str:
    """Sanitize a user-supplied song name into a safe .sng filename —
    same rule the desktop app uses when it slugs a default export name
    (constants.default_export_name), plus a path-traversal guard."""
    name = os.path.basename((name or "").strip()) or "untitled"
    name = _SAFE_NAME_RE.sub("", name).strip() or "untitled"
    if not name.lower().endswith(".sng"):
        name += ".sng"
    return name


class SongStore:
    """Reads/writes .sng files in one directory. No caching — each
    request re-reads from disk, so two browser tabs (or a script editing
    a file directly) never see stale data for long."""

    def __init__(self, directory: str):
        self.dir = directory
        os.makedirs(self.dir, exist_ok=True)

    def list_songs(self):
        out = []
        for fn in sorted(os.listdir(self.dir)):
            if not fn.lower().endswith(".sng"):
                continue
            try:
                doc = self.load(fn)
            except Exception:
                continue
            meta = doc.get("meta", {})
            out.append({
                "filename": fn,
                "title": meta.get("title", ""),
                "artist": meta.get("artist", ""),
                "sections": len(doc.get("sections", [])),
            })
        return out

    def _path(self, filename: str) -> str:
        filename = _safe_filename(filename)
        return os.path.join(self.dir, filename)

    def load(self, filename: str) -> dict:
        path = self._path(filename)
        with open(path, "r", encoding="utf-8") as f:
            raw = json.load(f)
        return model.migrate_document(raw)

    def save(self, filename: str, doc: dict) -> str:
        filename = _safe_filename(filename)
        doc = dict(doc)
        doc["app_version"] = APP_VERSION
        path = self._path(filename)
        with open(path, "w", encoding="utf-8") as f:
            json.dump(doc, f, indent=2)
        return filename

    def delete(self, filename: str):
        path = self._path(filename)
        if os.path.exists(path):
            os.remove(path)

    def exists(self, filename: str) -> bool:
        return os.path.exists(self._path(filename))


# ==============================================================================
#  Chart-line parse/preview helper (shared by /api/parse and /api/render) —
#  the browser equivalent of the Tkinter chart editor bar's live re-render.
# ==============================================================================

def _with_chart_lines(doc: dict) -> dict:
    """Return a shallow copy of `doc` with each section annotated with a
    transient `chart_line` string — the browser's chart-line <input>
    needs something to display; the desktop app instead keeps this in a
    live Tkinter Text widget, never in the saved document. Saving a doc
    back drops this field automatically (set_chart_items on the client
    rebuilds `items`, and the server never persists `chart_line`)."""
    out = dict(doc)
    out["sections"] = []
    for sec in doc.get("sections", []):
        sec = dict(sec)
        chart = songmap.chart_items(sec)
        # References display by the target's current name; the document
        # keeps ids (see songmap.display_ref_key).
        sec["chart_line"] = (grammar.unparse(songmap.display_items(chart, doc))
                             if chart else "")
        out["sections"].append(sec)
    return out


def _parse_chart_line(line: str, doc: dict = None, items=None):
    """Parse one chart line and render it the way the export will.

    `doc` is optional context: with it, a `=section` / riff reference on the
    line is expanded to the items it points at, so the section card's
    preview shows what will actually be played rather than the pointer.
    Without it (a bare parse), references render as themselves."""
    try:
        if items is None:
            items, annotation = grammar.parse(line)
        else:
            # Rendering what's *stored* rather than what's typed: a card that
            # references another section keeps the target's id, so it survives
            # a rename that the typed "=Old_Name" no longer resolves.
            items, annotation = list(items), None
        # Whatever the user typed — "=Interlude" or "=chorus1" — is stored
        # as the target's id, so the reference survives a later rename;
        # what comes back for display is spelled with the name again.
        if doc:
            items = songmap.canonicalise_refs(items, doc)
        display = render_mod.resolve_references(items, doc) if doc else items
        rows = render_mod.render_chart_rows(display) if display else []
        rendered = [r["text"] for r in rows]
        # A chart row is 1..N lines: an optional fret row, the symbol row,
        # then one line per string for any lick. Only the server knows
        # whether the *first* line is frets or symbols, so say so rather
        # than making the front end guess from the count. It has to be the
        # first row's own role, not "are there frets anywhere on the line":
        # a line broken with // whose opening block has no fret numbers
        # starts on symbols, and answering yes there handed the front end
        # a symbol row to paint as frets and a fret row to paint as
        # symbols — which is why a fret number went black after the first
        # block.
        has_fret_row = bool(rows) and rows[0]["role"] == render_mod.ROLE_FRET
        return {
            "ok": True,
            "items": items,
            "annotation": annotation,
            "unparsed": (grammar.unparse(songmap.display_items(items, doc))
                         if items else ""),
            "rendered": rendered,
            # Per row: what it is (fret / symbol / lick line) and any
            # stretch inside it that prints in its own colour — a rest, so
            # far. The browser colours from this rather than re-deriving it
            # from the text, which is how the preview and the PDF stay
            # the same chart.
            "roles": [r["role"] for r in rows],
            "spans": [r.get("spans") or [] for r in rows],
            "fret_row": has_fret_row,
        }
    except grammar.ParseError as exc:
        return {"ok": False, "error": str(exc)}


class Handler(BaseHTTPRequestHandler):
    server_version = f"SongNotationTool/{APP_VERSION}"
    store: SongStore = None  # set by main()

    # -- plumbing ---------------------------------------------------------

    def log_message(self, fmt, *args):  # quieter default logging
        pass

    def _send_json(self, obj, status=HTTPStatus.OK):
        body = json.dumps(obj).encode("utf-8")
        self.send_response(status)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Content-Length", str(len(body)))
        self.end_headers()
        self.wfile.write(body)

    def _send_bytes(self, body: bytes, content_type: str, filename: str = None):
        self.send_response(HTTPStatus.OK)
        self.send_header("Content-Type", content_type)
        self.send_header("Content-Length", str(len(body)))
        if filename:
            self.send_header("Content-Disposition",
                              f'attachment; filename="{filename}"')
        self.end_headers()
        self.wfile.write(body)

    def _send_error_json(self, status, message):
        self._send_json({"ok": False, "error": message}, status=status)

    def _read_json_body(self):
        length = int(self.headers.get("Content-Length", 0) or 0)
        if not length:
            return {}
        raw = self.rfile.read(length)
        return json.loads(raw.decode("utf-8"))

    def _serve_static(self, rel_path: str):
        if rel_path in ("", "/"):
            rel_path = "index.html"
        rel_path = rel_path.lstrip("/")
        full = os.path.normpath(os.path.join(WEB_DIR, rel_path))
        if not full.startswith(WEB_DIR):  # traversal guard
            self.send_error(HTTPStatus.FORBIDDEN)
            return
        if not os.path.isfile(full):
            self.send_error(HTTPStatus.NOT_FOUND)
            return
        ctype = mimetypes.guess_type(full)[0] or "application/octet-stream"
        with open(full, "rb") as f:
            self._send_bytes(f.read(), ctype)

    # -- routing ------------------------------------------------------------

    def do_GET(self):
        parsed = urlparse(self.path)
        path, qs = unquote(parsed.path), parse_qs(parsed.query)
        try:
            if path == "/api/meta":
                return self._send_json({
                    "app_version": APP_VERSION,
                    "songs_dir": self.store.dir,
                    "export_dir": userpaths.last_export_dir(),
                    "config_path": userpaths.config_path(),
                    "section_layouts": SECTION_LAYOUT_LABELS,
                    "lyrics_layouts": LYRICS_LAYOUT_LABELS,
                    "chord_sheets": CHORD_SHEET_LABELS,
                    "lick_refs": LICK_REF_LABELS,
                    "color_modes": COLOR_MODE_LABELS,
                    "pdf_scales": PDF_SCALE_LABELS,
                    "pdf_columns": PDF_COLUMN_LABELS,
                    "section_types": SECTION_TYPES,
                    "render_modes": RENDER_MODE_LABELS,
                    "instruments": INSTRUMENT_STRINGS,
                    "tab_beats_default": TAB_BEATS_DEFAULT,
                    "tab_beats_options": [8, 16, 32, 64],
                })
            if path == "/api/songs":
                return self._send_json(self.store.list_songs())

            m = re.match(r"^/api/songs/([^/]+)/export\.(txt|pdf)$", path)
            if m:
                filename, fmt = m.group(1), m.group(2)
                if not self.store.exists(filename):
                    return self._send_error_json(HTTPStatus.NOT_FOUND, "no such song")
                doc = self.store.load(filename)
                out_name = default_export_name(doc, fmt)
                if fmt == "txt":
                    body = "\n".join(export.build_song_lines(doc)).encode("utf-8")
                    return self._send_bytes(body, "text/plain; charset=utf-8", out_name)
                orient = (qs.get("orient", ["portrait"])[0])
                body = export.build_pdf(doc, orient=orient)
                return self._send_bytes(body, "application/pdf", out_name)

            m = re.match(r"^/api/songs/([^/]+)$", path)
            if m:
                filename = m.group(1)
                if not self.store.exists(filename):
                    return self._send_error_json(HTTPStatus.NOT_FOUND, "no such song")
                return self._send_json(_with_chart_lines(self.store.load(filename)))

            if path.startswith("/api/"):
                return self._send_error_json(HTTPStatus.NOT_FOUND, "unknown route")

            return self._serve_static(path)
        except Exception as exc:  # noqa: BLE001 — never crash the server thread
            self._send_error_json(HTTPStatus.INTERNAL_SERVER_ERROR, str(exc))

    def do_POST(self):
        parsed = urlparse(self.path)
        path = unquote(parsed.path)
        try:
            if path == "/api/quit":
                # "Save & Close" in the browser front end calls this after
                # its own save PUT completes. A short delay (matching the
                # reference project) lets this response actually reach the
                # browser before the process/window goes away.
                threading.Timer(0.5, _shutdown).start()
                return self._send_json({"ok": True, "message": "Shutting down…"})

            if path == "/api/songs":
                # New song: {name, title, artist, key, time, bpm}
                body = self._read_json_body()
                filename = _safe_filename(body.get("name") or body.get("title") or "untitled")
                if self.store.exists(filename):
                    return self._send_error_json(HTTPStatus.CONFLICT,
                                                   f"{filename} already exists")
                doc = model.new_document(
                    title=body.get("title", ""), artist=body.get("artist", ""),
                    key=body.get("key", ""), time=body.get("time", "4/4"),
                    bpm=body.get("bpm", ""))
                self.store.save(filename, doc)
                return self._send_json({"filename": filename, "doc": _with_chart_lines(doc)})

            if path == "/api/songs/example":
                # "Open example" — the built-in sample from Enhancement 1's
                # first-launch panel, so a new/returning user immediately
                # sees a real layout instead of an empty grid.
                filename = _safe_filename("Example Song")
                if self.store.exists(filename):
                    doc = self.store.load(filename)
                else:
                    doc = example_document()
                    self.store.save(filename, doc)
                return self._send_json({"filename": filename, "doc": _with_chart_lines(doc)})

            if path == "/api/parse":
                body = self._read_json_body()
                return self._send_json(
                    _parse_chart_line(body.get("line", ""), body.get("doc"),
                                      body.get("items")))

            if path == "/api/render":
                # Live preview pane (Enhancement 1): render an in-progress,
                # possibly-unsaved doc to TXT lines without touching disk.
                body = self._read_json_body()
                doc = body.get("doc") or {}
                doc = model.migrate_document(doc)
                instruments = body.get("instruments")
                lines = export.build_song_lines(
                    doc, instruments=set(instruments) if instruments else None)
                return self._send_json({"lines": lines})

            if path == "/api/lyrics/split":
                # Split the whole-song lyric sheet into blocks and propose
                # which section gets which — the front end shows the
                # proposal as one row per section to correct, and writes
                # the answer into the document itself. Shared with the
                # desktop app's Lyrics dialog so both make the same guess.
                body = self._read_json_body()
                doc = model.migrate_document(body.get("doc") or {})
                sheet = body.get("text")
                if sheet is None:
                    sheet = doc.get("lyrics_text") or ""
                blocks = lyrics_mod.split_blocks(sheet)
                sections = doc.get("sections", [])
                return self._send_json({
                    "blocks": blocks,
                    "suggested": lyrics_mod.suggest(
                        blocks, sections, lyrics_mod.repeated_blocks(sheet)),
                    "current": lyrics_mod.current_assignment(blocks, sections),
                })

            if path == "/api/lyrics/marked":
                # The sheet marked up with "=== Section ===" lines: say
                # which section each marked block belongs to, and which
                # names have no section yet. The front end offers to
                # create those; nothing is written here.
                body = self._read_json_body()
                doc = model.migrate_document(body.get("doc") or {})
                sheet = body.get("text")
                if sheet is None:
                    sheet = doc.get("lyrics_text") or ""
                segments = lyrics_mod.split_marked(sheet)
                return self._send_json({
                    "segments": segments,
                    "matched": lyrics_mod.match_sections(doc, segments),
                    "unmatched": lyrics_mod.unmatched_names(doc, sheet),
                })

            if path == "/api/chords/shape":
                # Parse one chord shape ("x32010") into a fret per string.
                # Parsing lives in chords.py so the browser, the desktop
                # app and the CLI all read a shape the same way.
                body = self._read_json_body()
                instrument = body.get("instrument") or chords_mod.DEFAULT_INSTRUMENT
                try:
                    frets = chords_mod.shape_from_text(body.get("text", ""), instrument)
                except chords_mod.ChordError as exc:
                    return self._send_json({"ok": False, "error": str(exc)})
                return self._send_json({
                    "ok": True, "frets": frets,
                    "strings": chords_mod.strings_for(instrument),
                    "text": chords_mod.shape_to_text(frets, instrument),
                })

            if path == "/api/chords/coverage":
                # Which chords the chart plays with no shape, and which
                # shapes it never plays — reported, never enforced.
                body = self._read_json_body()
                doc = model.migrate_document(body.get("doc") or {})
                cov = chords_mod.coverage(doc)
                cov["used"] = chords_mod.used_symbols(doc)
                return self._send_json(cov)

            if path == "/api/export.txt" or path == "/api/export.pdf":
                # Ad-hoc export of an unsaved doc straight from the editor.
                body = self._read_json_body()
                doc = model.migrate_document(body.get("doc") or {})
                out_name = default_export_name(doc, "txt" if path.endswith("txt") else "pdf")
                if path.endswith("txt"):
                    data = "\n".join(export.build_song_lines(doc)).encode("utf-8")
                    return self._send_bytes(data, "text/plain; charset=utf-8", out_name)
                orient = body.get("orient", "portrait")
                data = export.build_pdf(doc, orient=orient)
                return self._send_bytes(data, "application/pdf", out_name)

            return self._send_error_json(HTTPStatus.NOT_FOUND, "unknown route")
        except grammar.ParseError as exc:
            self._send_error_json(HTTPStatus.BAD_REQUEST, str(exc))
        except Exception as exc:  # noqa: BLE001
            self._send_error_json(HTTPStatus.INTERNAL_SERVER_ERROR, str(exc))

    def do_PUT(self):
        parsed = urlparse(self.path)
        m = re.match(r"^/api/songs/([^/]+)$", unquote(parsed.path))
        if not m:
            return self._send_error_json(HTTPStatus.NOT_FOUND, "unknown route")
        try:
            filename = m.group(1)
            doc = self._read_json_body()
            doc = model.migrate_document(doc)
            for sec in doc.get("sections", []):
                sec.pop("chart_line", None)
            saved_name = self.store.save(filename, doc)
            return self._send_json({"ok": True, "filename": saved_name})
        except Exception as exc:  # noqa: BLE001
            self._send_error_json(HTTPStatus.INTERNAL_SERVER_ERROR, str(exc))

    def do_DELETE(self):
        parsed = urlparse(self.path)
        m = re.match(r"^/api/songs/([^/]+)$", unquote(parsed.path))
        if not m:
            return self._send_error_json(HTTPStatus.NOT_FOUND, "unknown route")
        self.store.delete(m.group(1))
        self._send_json({"ok": True})


# ==============================================================================
#  Native window by default — re-exec into the project's own virtualenv.
#
#  pywebview is optional, but once it *is* installed the native window is
#  what you want every time, without having to remember to activate a venv
#  first. `python3 webserver.py` with the system interpreter would otherwise
#  silently fall back to a browser tab just because that interpreter can't
#  see ./.venv/lib/.../pywebview. So: if this interpreter has no pywebview
#  but a sibling virtualenv does, hand the whole process over to that
#  interpreter (os.execv — same PID, same argv, no wrapper script and no
#  environment variables for the user to set).
#
#  SNT_NO_REEXEC=1 disables it, and the child always carries that flag so a
#  broken venv can never produce an exec loop.
# ==============================================================================

_VENV_CANDIDATES = (".venv", "venv", "env")


def _same_path(a: str, b: str) -> bool:
    return os.path.realpath(a) == os.path.realpath(b)


def _venv_python(root: str):
    """
    Path to a sibling virtualenv interpreter that can import webview, or
    None. Cheap: only spawns a probe for venvs that actually exist.

    "Am I already running under this venv?" is answered with sys.prefix
    against the venv directory — NOT by comparing interpreter paths. On
    macOS (and most Linux venvs) .venv/bin/python3 is a symlink straight
    back to the base interpreter that created it, so `python3
    webserver.py` from a shell whose python3 IS that base interpreter
    makes realpath(.venv/bin/python3) == realpath(sys.executable) even
    though this process has none of the venv's site-packages. That false
    match is exactly the case this whole function exists to handle, so it
    must not be the thing that aborts it.

    Returns (path, None) on success, or (None, reason) so the caller can
    say something useful when a venv is present but unusable.
    """
    reason = None
    for name in _VENV_CANDIDATES:
        vroot = os.path.join(root, name)
        if not os.path.isdir(vroot):
            continue
        if _same_path(sys.prefix, vroot):
            return None, None       # genuinely already inside it
        for rel in (("bin", "python3"), ("bin", "python"), ("Scripts", "python.exe")):
            cand = os.path.join(vroot, *rel)
            if not os.path.isfile(cand):
                continue
            try:
                proc = subprocess.run([cand, "-c", "import webview"],
                                      stdout=subprocess.DEVNULL,
                                      stderr=subprocess.PIPE, timeout=30)
            except (OSError, subprocess.SubprocessError) as exc:
                reason = f"{os.path.relpath(cand, root)}: {exc}"
                continue
            if proc.returncode == 0:
                return cand, None
            err = (proc.stderr or b"").decode("utf-8", "replace").strip()
            last = err.splitlines()[-1] if err else f"exit {proc.returncode}"
            reason = f"{os.path.relpath(cand, root)} can't import webview — {last}"
            if os.environ.get("SNT_DEBUG_LAUNCH") and err:
                print(err)
            break                   # one interpreter per venv is enough
    return None, reason


def _maybe_reexec_into_venv(argv):
    """If the native window is wanted but this interpreter lacks pywebview,
    re-exec under a sibling virtualenv that has it. Returns only when there
    is nothing to do."""
    if os.environ.get("SNT_NO_REEXEC"):
        return
    if "--browser" in argv or "--no-open" in argv:
        return
    try:
        import webview  # noqa: F401
        return                      # already good — nothing to do
    except ImportError:
        pass

    root = os.path.dirname(os.path.abspath(__file__))
    py, reason = _venv_python(root)
    if not py:
        # A venv that exists but can't import webview is worth one line —
        # silently falling back to a browser tab is what made this
        # confusing in the first place.
        if reason:
            print(f"Note: {reason}")
            print("  (set SNT_DEBUG_LAUNCH=1 for the full traceback)")
        return

    os.environ["SNT_NO_REEXEC"] = "1"
    rel = os.path.relpath(py, root)
    print(f"Using {rel} (it has pywebview) for the native window.")
    try:
        os.execv(py, [py, os.path.abspath(__file__), *argv])
    except OSError as exc:          # exec failed — carry on in this process
        print(f"Could not switch interpreter ({exc}); continuing without it.")


def main(argv=None):
    _maybe_reexec_into_venv(list(sys.argv[1:] if argv is None else argv))
    parser = argparse.ArgumentParser(description=__doc__.strip().splitlines()[0])
    parser.add_argument("--dir", default=None,
                         help="Folder of .sng files to serve (default: your "
                              "configured songs folder, initially "
                              "~/Documents/Song Notation Tool)")
    parser.add_argument("--host", default="localhost",
                         help="Bind address (use 0.0.0.0 to reach it from "
                              "other devices on the network — default: localhost)")
    parser.add_argument("--port", type=int, default=8420)
    parser.add_argument("--browser", action="store_true",
                         help="Open in your default browser tab instead of a "
                              "native window (also used automatically when "
                              "pywebview isn't installed)")
    parser.add_argument("--no-open", action="store_true",
                         help="Don't open anything automatically — just start "
                              "the server (e.g. when only serving other "
                              "devices on the network)")
    args = parser.parse_args(argv)

    # Songs are the user's data, so they live under the user's home by
    # default — not in ./songs inside the checkout, where a `git clean` or
    # a re-clone would take them. --dir still overrides, per run, without
    # touching the stored setting.
    if args.dir:
        songs_dir = os.path.abspath(os.path.expanduser(args.dir))
    else:
        songs_dir = userpaths.songs_dir()
        legacy = os.path.join(os.path.dirname(os.path.abspath(__file__)), "songs")
        moved, skipped = userpaths.migrate_legacy_songs(legacy, songs_dir)
        if moved:
            print(f"Moved {len(moved)} song{'s' if len(moved) != 1 else ''} out of "
                  f"the source folder into {songs_dir}:")
            for name in moved:
                print(f"  {name}")
        if skipped:
            print(f"Left {len(skipped)} file(s) in {legacy} "
                  f"(a file of the same name was already in the songs folder):")
            for name in skipped:
                print(f"  {name}")

    Handler.store = SongStore(songs_dir)

    class ThreadingHTTPServer(socketserver.ThreadingMixIn, socketserver.TCPServer):
        allow_reuse_address = True
        daemon_threads = True

    httpd = ThreadingHTTPServer((args.host, args.port), Handler)
    _RUNTIME["httpd"] = httpd
    # 0.0.0.0 binds every interface but isn't itself a browsable address —
    # open localhost even when serving on 0.0.0.0 for other devices.
    open_host = "localhost" if args.host == "0.0.0.0" else args.host
    url = f"http://{open_host}:{args.port}"

    print(f"Song Notation Tool v{APP_VERSION}")
    print(f"Songs folder: {songs_dir}")
    if not args.dir:
        print(f"  (change it in the app, or edit {userpaths.config_path()})")

    use_webview = False
    pywebview_missing = False
    if not args.browser and not args.no_open:
        try:
            import webview  # pywebview — optional; falls back to a browser
            use_webview = True
        except ImportError:
            pywebview_missing = True

    if use_webview:
        # ── Native window mode: a normal titled OS window — not a browser
        # tab (no address bar, no tabs), but not frameless either. This
        # started out frameless; reverted after testing found that a
        # borderless pywebview window loses native resize-by-edge and
        # maximize on macOS, with no reliable custom replacement pywebview
        # exposes for either. A titlebar's own maximize/resize/close
        # buttons are the correct fix, not more code — matches how the
        # sibling qlc-plus-swiss-knife-tool-script project's window
        # actually works too. resizable=True is pywebview's own default;
        # named explicitly here since it's the whole point.
        # pywebview needs the main thread for its OS event loop (required
        # on macOS in particular), so the HTTP server runs in a background
        # thread instead of the usual serve_forever() on the main thread.
        server_thread = threading.Thread(target=httpd.serve_forever, daemon=True)
        server_thread.start()
        print(f"Opening {url} in a native window.  Close the window (or "
              f"use the in-app Save & Close button) to quit.")
        window = webview.create_window(
            "Song Notation Tool", url,
            width=1180, height=820, min_size=(900, 600),
            resizable=True, confirm_close=True,
            js_api=_JsApi(),
        )
        _RUNTIME["webview_window"] = window
        webview.start()
        httpd.shutdown()
        httpd.server_close()
        print("Window closed — bye!")
        return

    # ── Browser mode (also the fallback when pywebview isn't installed):
    # open the default browser shortly after the server starts accepting
    # connections, then block on serve_forever() as before.
    if pywebview_missing:
        print("Note: pywebview isn't installed for any interpreter this "
              "script can find, so it opened in your browser instead of a "
              "native app window.")
        print("  Set it up once, and every later launch is a native window:")
        print("    python3 -m venv .venv")
        print("    .venv/bin/python3 -m pip install -r requirements-optional.txt")
        print("  After that, plain `python3 webserver.py` finds it by itself.")
        print("Other options: --browser (always use a browser tab), "
              "--no-open (don't open anything automatically)")
    if not args.no_open:
        threading.Timer(1.0, lambda: webbrowser.open(url)).start()
        print(f"Opening {url} in your browser.  Ctrl+C to stop.")
    else:
        print(f"Open {url} in a browser.  Ctrl+C to stop.")
    try:
        httpd.serve_forever()
    except KeyboardInterrupt:
        pass
    finally:
        httpd.server_close()


if __name__ == "__main__":
    main()
