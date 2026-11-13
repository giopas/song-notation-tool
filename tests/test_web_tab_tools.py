"""v0.36 in the browser app: the tab grid's tools (bar lines, selection,
copy and paste, repeat, name), copies that bring their parts, the PDF
preview, and older songs whose tab moves into the grid when opened.

The clicks themselves run in the browser and were checked there (headless
Chromium); these tests keep the page, the script and the server in step.
"""
import json
import os
import re
import sys
import threading
import urllib.error
import urllib.request

import pytest

sys.path.insert(0, os.path.join(os.path.dirname(__file__), ".."))

import grammar  # noqa: E402
import model  # noqa: E402

HERE = os.path.dirname(__file__)
INDEX = open(os.path.join(HERE, "..", "web", "index.html"), encoding="utf-8").read()
APP = open(os.path.join(HERE, "..", "web", "app.js"), encoding="utf-8").read()
CSS = open(os.path.join(HERE, "..", "web", "style.css"), encoding="utf-8").read()


def test_the_grid_has_its_tools():
    for fn in ("function ensureGridTools(", "function toggleBar(", "function barsEvery(",
               "function splitMeasureAt(", "function repeatTabRange(",
               "function copyTabRange(", "function pasteTab(", "function clearTabRange(",
               "function openMeasureMenu(", "function tabNameProblem(",
               "function renameLickRefs("):
        assert fn in APP, fn
    for cls in ("tab-copy", "tab-paste", "tab-repeat", "tab-bar", "tab-name"):
        assert f'class="btn subtle {cls}"' in APP or f'class="{cls}"' in APP, cls
    for cls in (".tab-bar-before", ".tab-rep-open", ".tab-rep-close", ".tab-sel",
                ".tab-bar-hit", ".tab-measure-menu"):
        assert cls in CSS, cls


def test_typing_a_bar_in_a_cell_makes_a_bar_line():
    body = APP[APP.index("function onTabCellKey("):]
    body = body[:body.index("\n}\n")]
    assert 'e.key === "|"' in body
    assert "toggleBar(sec, c.mi, c.bi + 1)" in body


def test_copy_and_paste_keys_work_on_a_selection():
    body = APP[APP.index("function onTabCellKey("):]
    body = body[:body.index("\n}\n")]
    assert "copyTabRange(sec, r[0], r[1])" in body
    assert "pasteTab(sec, c.g, TAB_CLIP)" in body
    assert 'e.shiftKey && (e.key === "ArrowRight"' in body


def test_a_named_tab_is_offered_as_a_lick():
    body = APP[APP.index("function namedLicks()"):]
    body = body[:body.index("\n}\n")]
    assert "tab_name" in body


def test_the_name_rules_match_the_server():
    import songmap
    doc = model.new_document("T", "B")
    for name in ("Riff1", "Hook", "A1", "Am", "two words", "mf", "pm"):
        server_ok = songmap.check_tab_name(doc, name) == ""
        # The script's first two checks are the server's regular expressions.
        script_ok = bool(re.match(r"^[A-Za-z][A-Za-z0-9_]*$", name)) and \
            not re.match(r"^\d{0,2}[A-Ga-g]([#b])?", name) and \
            name not in ("pp", "p", "mp", "mf", "f", "ff") and \
            name.lower() not in ("pm", "cresc", "dim")
        assert server_ok == script_ok, name
    assert "/^\\d{0,2}[A-Ga-g]([#b])?/" in APP


def test_a_reference_copy_brings_every_part():
    body = APP[APP.index("function duplicateSection(id, at)"):]
    body = body[:body.index("\n}\n")]
    assert "newSec.parts = source.parts.map(" in body
    assert "ref:" in body


def test_the_menu_has_an_independent_copy():
    tpl = INDEX[INDEX.index('<template id="section-template">'):]
    assert 'class="sec-dup-copy"' in tpl
    assert 'openCopyPicker(sec.id, node.querySelector(".sec-more"), "copy")' in APP
    assert "function duplicateSectionCopy(id, at)" in APP


def test_a_part_reference_is_shown_as_one():
    assert "if (part.ref) { renderPartRef(node, sec, box, part); return; }" in APP


def test_the_preview_can_show_the_pdf():
    pane = INDEX[INDEX.index('id="preview-pane"'):]
    pane = pane[:pane.index("</section>")]
    for el in ('id="preview-mode"', 'data-preview="pdf"', 'id="preview-pdf-frame"',
               'id="preview-orient"', 'id="btn-preview-open"'):
        assert el in pane, el
    assert '"/api/preview.pdf"' in APP
    assert "wirePreviewMode();" in APP


def test_an_upgraded_song_says_so_and_is_not_saved_yet():
    body = APP[APP.index("async function openSong(filename)"):]
    body = body[:body.index("\n}\n")]
    assert "doc._upgraded" in body and "delete doc._upgraded" in body
    assert "savedSnapshot = null" in body


# ── the HTTP routes ──────────────────────────────────────────────────────

@pytest.fixture()
def server(tmp_path):
    import webserver
    old = getattr(webserver.Handler, "store", None)
    webserver.Handler.store = webserver.SongStore(str(tmp_path / "songs"))
    httpd = webserver._ThreadingHTTPServer(("127.0.0.1", 0), webserver.Handler)
    threading.Thread(target=httpd.serve_forever, daemon=True).start()
    base = f"http://127.0.0.1:{httpd.server_address[1]}"

    def call(path, body=None):
        data = None if body is None else json.dumps(body).encode()
        req = urllib.request.Request(base + path, data=data,
                                     headers={"Content-Type": "application/json"})
        try:
            with urllib.request.urlopen(req, timeout=20) as r:
                raw = r.read()
                if r.headers.get("Content-Type", "").startswith("application/json"):
                    return r.status, json.loads(raw), r.headers
                return r.status, raw, r.headers
        except urllib.error.HTTPError as e:
            return e.code, json.loads(e.read()), e.headers

    yield call, tmp_path / "songs"
    httpd.shutdown()
    httpd.server_close()
    webserver.Handler.store = old


def _old_song(folder):
    doc = model.new_document("Old One", "The Testers")
    sec = model.new_section("s1", "Intro", instrument="Bass (4-string)", render="free",
                            free_text="G|--5--7--|\nD|--------|\nA|--------|\nE|-3------|")
    doc["sections"].append(sec)
    lick = model.new_section("s2", "Riff", instrument="Bass (4-string)")
    lick["items"], _ = grammar.parse("{G 5 7 5 | D - - 3}")
    doc["sections"].append(lick)
    os.makedirs(folder, exist_ok=True)
    path = os.path.join(folder, "The Testers - Old One.sng")
    with open(path, "w", encoding="utf-8") as f:
        json.dump(doc, f)
    return path


def test_opening_an_old_song_moves_its_tab_into_the_grid(server):
    call, folder = server
    path = _old_song(folder)
    before = open(path, encoding="utf-8").read()
    status, doc, _ = call("/api/songs/The%20Testers%20-%20Old%20One.sng")
    assert status == 200
    assert [u["section"] for u in doc["_upgraded"]] == ["Intro", "Riff"]
    assert [s["render"] for s in doc["sections"]] == ["tab", "tab"]
    assert doc["tab_layout"] == 2
    # Only opened, not saved: the file is as it was.
    assert open(path, encoding="utf-8").read() == before


def test_a_new_song_is_never_upgraded(server):
    call, folder = server
    status, res, _ = call("/api/songs", {"title": "Fresh", "artist": "The Testers"})
    assert res["doc"]["tab_layout"] == 2


def test_the_pdf_preview_is_served_inline(server):
    call, folder = server
    doc = model.new_document("Preview Me", "The Testers")
    sec = model.new_section("s1", "Verse")
    sec["items"], _ = grammar.parse("5A 7D")
    doc["sections"].append(sec)
    status, res, _ = call("/api/preview.pdf", {"doc": doc, "orient": "landscape"})
    assert status == 200 and res["ok"]
    assert re.fullmatch(r"/api/preview/[0-9a-f]{16}\.pdf", res["url"])
    status, data, headers = call(res["url"])
    assert status == 200 and data.startswith(b"%PDF")
    assert headers["Content-Type"] == "application/pdf"
    assert headers["Content-Disposition"].startswith("inline")
    status, _, _ = call("/api/preview/0123456789abcdef.pdf")
    assert status == 404


def test_the_pdf_preview_reports_a_song_that_cannot_print(server):
    call, folder = server
    doc = model.new_document("Broken", "The Testers")
    sec = model.new_section("s1", "Verse")
    sec["items"] = [{"kind": "group", "repeat": 0, "items": []}]
    doc["sections"].append(sec)
    status, res, _ = call("/api/preview.pdf", {"doc": doc})
    assert status in (200, 422)
    if status == 422:
        assert res["errors"]


def test_only_a_few_previews_are_kept():
    import webserver
    webserver._PREVIEW_PDFS.clear()
    tokens = [webserver._remember_preview(b"%PDF-" + bytes([i])) for i in range(10)]
    assert len(webserver._PREVIEW_PDFS) == webserver._PREVIEW_KEEP
    assert tokens[-1] in webserver._PREVIEW_PDFS and tokens[0] not in webserver._PREVIEW_PDFS
