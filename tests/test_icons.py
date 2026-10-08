"""The app icon: every file the app, the web front end and the packages
use is there, in the format its name promises, and is referenced where it
is shown."""
import os
import struct

ROOT = os.path.abspath(os.path.join(os.path.dirname(__file__), ".."))
PNG_SIG = b"\x89PNG\r\n\x1a\n"


def _read(rel):
    with open(os.path.join(ROOT, rel), "rb") as f:
        return f.read()


def _png_size(data):
    assert data[:8] == PNG_SIG
    return struct.unpack(">II", data[16:24])


def test_png_icons_exist_at_their_stated_sizes():
    for rel, size in [("web/logo/icon-256.png", 256), ("web/logo/icon-180.png", 180),
                      ("web/logo/favicon.png", 64), ("packaging/icons/icon-512.png", 512),
                      ("packaging/icons/icon-1024.png", 1024)]:
        assert _png_size(_read(rel)) == (size, size), rel


def test_svg_source_has_no_text_so_it_does_not_depend_on_fonts():
    svg = _read("web/logo/icon.svg").decode("utf-8")
    assert svg.lstrip().startswith("<svg") and "<text" not in svg


def test_icns_and_ico_have_the_right_headers():
    icns = _read("packaging/icons/icon.icns")
    assert icns[:4] == b"icns" and struct.unpack(">I", icns[4:8])[0] == len(icns)
    ico = _read("packaging/icons/icon.ico")
    reserved, kind, count = struct.unpack("<HHH", ico[:6])
    assert (reserved, kind) == (0, 1) and count >= 5


def test_web_front_end_uses_the_icon():
    html = _read("web/index.html").decode("utf-8")
    assert 'href="/logo/favicon.png"' in html
    assert 'src="/logo/icon.svg"' in html


def test_webserver_passes_the_icon_only_when_pywebview_accepts_it():
    import webserver

    class New:
        @staticmethod
        def start(func=None, icon=None):
            pass

    class Old:
        @staticmethod
        def start(func=None):
            pass

    assert webserver._webview_icon_kwargs(New) == {"icon": webserver.ICON_PNG}
    assert webserver._webview_icon_kwargs(Old) == {}
