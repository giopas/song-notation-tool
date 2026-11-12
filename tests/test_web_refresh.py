"""The v0.34.0 browser front end: what the page and the script must keep.

The look itself is checked in a browser (tools/shots.py and a live run).
These tests keep the page, the script and the server in step on the
parts a later change could quietly undo: every dialog closes with Escape,
the song tools exist to be switched off, the render-mode buttons are all
there, the font is served, and a measure made longer or shorter still
prints and exports.
"""
import os
import re
import threading
import urllib.request

import pytest

import export
import model
import render

HERE = os.path.dirname(__file__)
ROOT = os.path.join(HERE, "..")
INDEX = open(os.path.join(ROOT, "web", "index.html"), encoding="utf-8").read()
APP = open(os.path.join(ROOT, "web", "app.js"), encoding="utf-8").read()
CSS = open(os.path.join(ROOT, "web", "style.css"), encoding="utf-8").read()


def _template():
    start = INDEX.index('<template id="section-template">')
    return INDEX[start:INDEX.index("</template>", start)]


def test_every_dialog_names_the_button_escape_presses():
    backdrops = re.findall(r'<div id="([\w-]+)" class="modal-backdrop hidden"([^>]*)>', INDEX)
    assert len(backdrops) >= 10
    for bid, attrs in backdrops:
        m = re.search(r'data-close="([\w-]+)"', attrs)
        assert m, f"{bid} has no data-close"
        assert f'id="{m.group(1)}"' in INDEX, f"{bid}: no button {m.group(1)}"


def test_every_dialog_is_a_labelled_modal_dialog():
    for m in re.finditer(r'<div[^>]*class="modal( modal-wide)?"[^>]*>', INDEX):
        tag = m.group(0)
        assert 'role="dialog"' in tag and 'aria-modal="true"' in tag, tag
        label = re.search(r'aria-labelledby="([\w-]+)"', tag).group(1)
        assert f'id="{label}"' in INDEX, label


def test_ids_are_unique():
    ids = re.findall(r'\sid="([^"]+)"', INDEX)
    assert len(ids) == len(set(ids))


def test_song_tools_can_be_switched_off_until_a_song_is_open():
    for bid in ("btn-transpose", "btn-lyrics", "btn-riffs", "btn-chords",
                "btn-preview-toggle", "btn-save", "btn-export"):
        tag = re.search(rf'<button id="{bid}"[^>]*>', INDEX).group(0)
        assert "song-tool" in tag, bid
    assert "setSongToolsEnabled(false)" in APP
    assert "setSongToolsEnabled(true)" in APP


def test_the_top_bar_has_no_emoji():
    top = INDEX[INDEX.index('<header id="topbar">'):INDEX.index("</header>")]
    assert not re.search("[\U0001F300-\U0001FAFF]", top)


def test_icon_only_buttons_have_a_name():
    for m in re.finditer(r'<button[^>]*class="[^"]*icon-btn[^"]*"[^>]*>', INDEX):
        assert "aria-label=" in m.group(0), m.group(0)


def test_the_four_render_modes_are_buttons_in_the_section_header():
    tpl = _template()
    modes = re.findall(r'data-mode="(\w+)"', tpl)
    assert modes == ["chart", "tab", "both", "free"]


def test_the_insert_buttons_are_grouped_and_only_under_the_active_card():
    tpl = _template()
    assert [g for g in re.findall(r'<span class="ins-label">(\w+)</span>', tpl)] == \
        ["Lines", "Repeats", "Licks", "Marks"]
    assert ".section-card.is-active .sec-insert-strip" in CSS


def test_the_error_sits_under_the_chart_line():
    tpl = _template()
    assert tpl.index("sec-chart-line") < tpl.index("sec-error") < tpl.index("sec-insert-strip")


def test_the_size_menu_names_the_default():
    tpl = _template()
    assert 'data-size="normal"' in tpl and ">Size<" not in tpl


def test_grey_text_reads_on_every_surface():
    """--lbl-gray against the dark surfaces, at 4.5:1 or better."""
    def lum(h):
        c = [int(h[i:i + 2], 16) / 255 for i in (1, 3, 5)]
        c = [x / 12.92 if x <= 0.03928 else ((x + 0.055) / 1.055) ** 2.4 for x in c]
        return 0.2126 * c[0] + 0.7152 * c[1] + 0.0722 * c[2]

    def ratio(a, b):
        hi, lo = sorted([lum(a), lum(b)], reverse=True)
        return (hi + 0.05) / (lo + 0.05)

    root = CSS[CSS.index(":root {"):CSS.index("}", CSS.index(":root {"))]
    var = dict(re.findall(r"--([\w-]+):\s*(#[0-9a-fA-F]{6})", root))
    for surface in ("bg", "card-bg", "surface", "input-bg", "header-bg"):
        assert ratio(var["lbl-gray"], var[surface]) >= 4.5, surface
    assert ratio(var["accent-fg"], var["accent"]) >= 4.5


def test_the_font_ships_with_its_licence():
    fonts = os.path.join(ROOT, "web", "fonts")
    assert os.path.exists(os.path.join(fonts, "OFL.txt"))
    for name in ("JetBrainsMono-Regular.woff2", "JetBrainsMono-Bold.woff2"):
        with open(os.path.join(fonts, name), "rb") as f:
            assert f.read(4) == b"wOF2", name
        assert f"/fonts/{name}" in CSS


def test_the_server_sends_the_font_as_a_font(tmp_path):
    import webserver
    old = getattr(webserver.Handler, "store", None)
    webserver.Handler.store = webserver.SongStore(str(tmp_path / "songs"))
    httpd = webserver._ThreadingHTTPServer(("127.0.0.1", 0), webserver.Handler)
    t = threading.Thread(target=httpd.serve_forever, daemon=True)
    t.start()
    try:
        port = httpd.server_address[1]
        with urllib.request.urlopen(
                f"http://127.0.0.1:{port}/fonts/JetBrainsMono-Regular.woff2") as r:
            assert r.headers["Content-Type"] == "font/woff2"
    finally:
        httpd.shutdown()
        webserver.Handler.store = old


def test_the_woff2_type_is_registered():
    import mimetypes
    import webserver  # noqa: F401  (registers the type on import)
    assert mimetypes.guess_type("x.woff2")[0] == "font/woff2"


def test_the_script_can_make_a_measure_longer_and_shorter():
    assert re.search(r"function resizeMeasure\(node, sec, body, mi, delta\)", APP)
    assert re.search(r"function emptyTail\(m, strings\)", APP)
    # Only empty beats are ever taken off.
    body = APP[APP.index("function resizeMeasure"):APP.index("function commitTabCell")]
    assert "emptyTail(m, strings)" in body


def _song_with_measures(*measures):
    doc = model.new_document("Odd Measures", "Invented Band")
    sec = model.new_section("s1", "Riff", "Verse", "Bass (4-string)", render="tab")
    sec["items"] = [model.make_measure(b, s) for b, s in measures]
    doc["sections"] = [sec]
    return doc


def test_a_longer_measure_prints_every_beat():
    """+4 on the second measure: 8 beats, then 12. Both print, the second
    with all twelve, and the song still passes the export check."""
    doc = _song_with_measures(
        (8, {"G": "", "D": "", "A": "- - 5 - - - 5 -", "E": "3 - - - 3 - - -"}),
        (12, {"G": "- - - - - - - - - - - -", "D": "- - - - - 7 - - - - - -",
              "A": "- - 5 - - - - - - 5 - -", "E": "3 - - - 5 - - - - - 3 -"}))
    assert export.check_document(doc)["errors"] == []
    lines = render.tab_grid_lines(doc["sections"][0]["items"], "Bass (4-string)")
    e_row = next(line for line in lines if line.startswith("E|"))
    measures = [m.split() for m in e_row[2:].split("|") if m.strip()]
    assert [len(m) for m in measures] == [8, 12]
    txt, _ = export.prepare_export(doc, "txt")
    pdf, _ = export.prepare_export(doc, "pdf")
    assert b"7" in txt and pdf[:5] == b"%PDF-"
