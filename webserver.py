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
    _RUNTIME["httpd"] = httpd
    # 0.0.0.0 binds every interface but isn't itself a browsable address —
    # open localhost even when serving on 0.0.0.0 for other devices.
    open_host = "localhost" if args.host == "0.0.0.0" else args.host
    url = f"http://{open_host}:{args.port}"

    print(f"Song Notation Tool v{APP_VERSION} — serving {os.path.abspath(args.dir)}")

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
