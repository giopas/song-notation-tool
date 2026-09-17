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
                                    # a native chrome-less window if
                                    # pywebview is installed, else
                                    # your default browser
    ./webserver.py --dir mysongs --port 9000
    ./webserver.py --browser       # force a normal browser tab, even
                                    # if pywebview is installed
    ./webserver.py --no-open       # just start the server, open
                                    # nothing (e.g. when only serving
                                    # other devices on the network)
    ./webserver.py --host 0.0.0.0  # reachable from other devices on
                                    # the LAN

pywebview (`pip install pywebview`) is entirely optional — like the
sibling qlc-plus-swiss-knife-tool-script project, its native-window
mode is used automatically when it's importable and falls straight
back to opening your default browser tab when it isn't. Nothing here
requires it to run.

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
import threading
import webbrowser
from http import HTTPStatus
from http.server import BaseHTTPRequestHandler
from urllib.parse import urlparse, parse_qs, unquote

import export
import grammar
import model
import songmap
import transpose
import render as render_mod
from examples import example_document
from constants import (
    APP_VERSION, INSTRUMENT_STRINGS, SECTION_TYPES, RENDER_MODE_LABELS,
    TAB_BEATS_DEFAULT, default_export_name,
)

WEB_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "web")

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
        sec["chart_line"] = grammar.unparse(chart) if chart else ""
        out["sections"].append(sec)
    return out


def _parse_chart_line(line: str):
    try:
        items, annotation = grammar.parse(line)
        rendered = render_mod.render_chart_row(items) if items else []
        return {
            "ok": True,
            "items": items,
            "annotation": annotation,
            "unparsed": grammar.unparse(items) if items else "",
            "rendered": rendered,
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
                    "section_types": SECTION_TYPES,
                    "render_modes": RENDER_MODE_LABELS,
                    "instruments": INSTRUMENT_STRINGS,
                    "tab_beats_default": TAB_BEATS_DEFAULT,
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
                return self._send_json(_parse_chart_line(body.get("line", "")))

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


def main(argv=None):
    parser = argparse.ArgumentParser(description=__doc__.strip().splitlines()[0])
    parser.add_argument("--dir", default="songs",
                         help="Folder of .sng files to serve (default: ./songs)")
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

    Handler.store = SongStore(args.dir)

    class ThreadingHTTPServer(socketserver.ThreadingMixIn, socketserver.TCPServer):
        allow_reuse_address = True
        daemon_threads = True

    httpd = ThreadingHTTPServer((args.host, args.port), Handler)
    # 0.0.0.0 binds every interface but isn't itself a browsable address —
    # open localhost even when serving on 0.0.0.0 for other devices.
    open_host = "localhost" if args.host == "0.0.0.0" else args.host
    url = f"http://{open_host}:{args.port}"

    print(f"Song Notation Tool v{APP_VERSION} — serving {os.path.abspath(args.dir)}")

    use_webview = False
    if not args.browser and not args.no_open:
        try:
            import webview  # pywebview — optional; falls back to a browser
            use_webview = True
        except ImportError:
            pass

    if use_webview:
        # ── Native window mode: a chrome-less window (no address bar or
        # tabs), like the sibling qlc-plus-swiss-knife-tool-script project.
        # pywebview needs the main thread for its OS event loop (required
        # on macOS in particular), so the HTTP server runs in a background
        # thread instead of the usual serve_forever() on the main thread.
        server_thread = threading.Thread(target=httpd.serve_forever, daemon=True)
        server_thread.start()
        print(f"Opening {url} in a native window.  Close the window to quit.")
        webview.create_window(
            "Song Notation Tool", url,
            width=1180, height=820, min_size=(900, 600),
        )
        webview.start()
        httpd.shutdown()
        httpd.server_close()
        return

    # ── Browser mode (also the fallback when pywebview isn't installed):
    # open the default browser shortly after the server starts accepting
    # connections, then block on serve_forever() as before.
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
