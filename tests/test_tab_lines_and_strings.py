"""v0.37: a repeat whose start or end moves to another measure, a tab grid
split into separate tabs on lines of their own, and the Layout setting that
prints the empty strings of a tab too."""
import json
import os
import shutil
import subprocess
import sys

import pytest

sys.path.insert(0, os.path.join(os.path.dirname(__file__), ".."))

import export  # noqa: E402
import model  # noqa: E402
import render  # noqa: E402

HERE = os.path.dirname(__file__)
APP = open(os.path.join(HERE, "..", "web", "app.js"), encoding="utf-8").read()
INDEX = open(os.path.join(HERE, "..", "web", "index.html"), encoding="utf-8").read()
CSS = open(os.path.join(HERE, "..", "web", "style.css"), encoding="utf-8").read()

BASS = "Bass (4-string)"


def measure(beats, **strings):
    return {"kind": "measure", "beats": beats,
            "strings": {st: strings.get(st, "") for st in ("G", "D", "A", "E")}}


def tab_song(*measures, **doc_keys):
    doc = model.new_document("Tab Song", "The Testers")
    doc.update(doc_keys)
    sec = model.new_section("intro", "Intro", instrument=BASS, render="tab")
    sec["items"] = list(measures)
    doc["sections"].append(sec)
    return doc


def e_rows(lines):
    return [ln for ln in lines if ln.startswith("E|")]


# ── separate tabs on lines of their own ─────────────────────────────────

def test_a_grid_with_no_marks_is_one_tab():
    ms = [measure(4) for _ in range(3)]
    assert render.grid_segments(ms) == [[0, 1, 2]]
    assert render.grid_rows(ms, 2) == [[0, 1], [2]]


def test_a_new_line_mark_starts_a_new_tab():
    ms = [measure(4) for _ in range(4)]
    ms[2]["new_line"] = True
    assert render.grid_segments(ms) == [[0, 1], [2, 3]]
    # A new tab starts a new line even when the measures would fit on one.
    assert render.grid_rows(ms, 10) == [[0, 1], [2, 3]]
    # Each tab still wraps on its own.
    assert render.grid_rows(ms, 1) == [[0], [1], [2], [3]]


def test_a_new_line_mark_on_the_first_measure_changes_nothing():
    ms = [measure(4), measure(4)]
    ms[0]["new_line"] = True
    assert render.grid_segments(ms) == [[0, 1]]


def test_the_txt_prints_two_tabs_on_two_lines():
    a, b = measure(4, E="3 - - -"), measure(4, E="5 - - -")
    b["new_line"] = True
    lines = render.tab_grid_lines([a, b], BASS)
    rows = e_rows(lines)
    assert len(rows) == 2
    assert "3" in rows[0] and "5" not in rows[0]
    assert "5" in rows[1]
    # The second tab keeps its measure number.
    assert any(ln.strip().startswith("M2") for ln in lines)
    text = "\n".join(export.build_song_lines(tab_song(a, b)))
    assert len(e_rows(text.splitlines())) == 2


def test_two_tabs_cost_two_lines_in_the_page_estimate():
    a, b = measure(4, E="3 - - -"), measure(4, E="5 - - -")
    one = tab_song(a, b)["sections"][0]
    b2 = dict(b, new_line=True)
    two = tab_song(a, b2)["sections"][0]
    strings = ["E"]
    assert (render.estimate_section_lines(two, strings)
            == render.estimate_section_lines(one, strings) + 1)


def test_the_pdf_builds_with_two_tabs():
    a, b = measure(4, E="3 - - -"), measure(4, E="5 - - -")
    b["new_line"] = True
    pdf = export.build_pdf(tab_song(a, b))
    assert pdf.startswith(b"%PDF")


def test_a_part_prints_two_tabs_too():
    a, b = measure(4, E="3 - - -"), measure(4, E="5 - - -")
    b["new_line"] = True
    doc = tab_song(measure(4, E="1 - - -"))
    doc["sections"][0]["parts"] = [{"id": "p1", "instrument": BASS, "render": "tab",
                                    "items": [a, b], "chart_line": ""}]
    rows = render.printed_part_rows(doc, doc["sections"][0])
    assert len([r for r in rows if r["text"].lstrip().startswith("E|")]) == 2


# ── empty strings ───────────────────────────────────────────────────────

def test_empty_strings_are_left_out_by_default():
    lines = render.tab_grid_lines([measure(4, E="3 - - -")], BASS)
    assert [ln[:2] for ln in lines if "|" in ln[:3]] == ["E|"]


def test_all_strings_prints_the_empty_ones_too():
    doc = {"tab_strings": "all"}
    lines = render.tab_grid_lines([measure(4, E="3 - - -")], BASS, doc=doc)
    labels = [ln[:2] for ln in lines if "|" in ln[:3]]
    assert labels == ["G|", "D|", "A|", "E|"]
    g_row = [ln for ln in lines if ln.startswith("G|")][0]
    assert set(g_row[2:].replace("|", "").split()) == {"-"}


def test_the_song_setting_reaches_the_txt_and_the_pdf():
    m = measure(4, E="3 - - -")
    used = "\n".join(export.build_song_lines(tab_song(m)))
    every = "\n".join(export.build_song_lines(tab_song(m, tab_strings="all")))
    assert "G|" not in used and "G|" in every
    assert len(export.build_pdf(tab_song(m, tab_strings="all"))) > \
        len(export.build_pdf(tab_song(m)))


def test_the_setting_is_offered_in_both_apps():
    import constants
    assert set(constants.TAB_STRINGS_LABELS) == {"used", "all"}
    assert 'id="meta-tab-strings"' in INDEX
    assert '["meta-tab-strings", "tab_strings", META.tab_strings, "used"]' in APP
    src = open(os.path.join(HERE, "..", "song_writer.py"), encoding="utf-8").read()
    assert '("tab_strings", "Tab strings", TAB_STRINGS_LABELS, "used"' in src
    ws = open(os.path.join(HERE, "..", "webserver.py"), encoding="utf-8").read()
    assert '"tab_strings": TAB_STRINGS_LABELS' in ws


# ── the browser app ─────────────────────────────────────────────────────

def test_the_grid_offers_the_new_tools():
    for fn in ("function addTabMeasure(", "function gridSegments(",
               "function repeatPairs(", "function moveRepeatStart(",
               "function moveRepeatEnd("):
        assert fn in APP, fn
    assert "Move the repeat start |: here" in APP
    assert "Move the repeat end :| here" in APP
    assert "Start a new tab line here" in APP
    assert 'class="btn subtle tab-add-line"' in INDEX
    assert "part-add-line" in APP
    assert ".tab-grid-next" in CSS


def _js(fn_names):
    out = []
    for name in fn_names:
        start = APP.index(f"function {name}(")
        out.append(APP[start:APP.index("\n}\n", start) + 3])
    return "\n".join(out)


NODE = shutil.which("node")


@pytest.mark.skipif(not NODE, reason="node is not installed")
def test_moving_a_repeat_in_the_browser_logic():
    script = _js(["repeatTimes", "gridSegments", "repeatPairs", "repeatToStartAt",
                  "repeatToEndAt", "moveRepeatStart", "moveRepeatEnd"]) + r"""
const mk = (n) => Array.from({ length: n }, () => ({ beats: 4 }));
const out = {};
// A repeat over the whole tab, from M1 to M6: start it at M4 instead.
let ms = mk(6); ms[0].repeat_open = true; ms[5].repeat_close = 3;
out.start = moveRepeatStart(ms, 3);
out.afterStart = repeatPairs(ms);
out.countKept = ms[5].repeat_close;
// Now end it at M5: the count goes with it.
out.end = moveRepeatEnd(ms, 4);
out.afterEnd = repeatPairs(ms);
out.countMoved = ms[4].repeat_close;
// A :| with no |: repeats from the start; its start can still move.
ms = mk(4); ms[3].repeat_close = 2;
out.implicit = repeatPairs(ms);
moveRepeatStart(ms, 1);
out.implicitMoved = repeatPairs(ms);
// Nothing to move: no repeat at all.
out.none = [repeatToStartAt(mk(3), 1), repeatToEndAt(mk(3), 1)];
// Two repeats: moving the second one's start leaves the first alone.
ms = mk(6); ms[0].repeat_open = true; ms[1].repeat_close = 2;
ms[2].repeat_open = true; ms[5].repeat_close = 4;
moveRepeatStart(ms, 4);
out.two = repeatPairs(ms);
// Segments: a measure marked new_line starts the next tab.
ms = mk(4); ms[2].new_line = true; ms[0].new_line = true;
out.segments = gridSegments(ms);
console.log(JSON.stringify(out));
"""
    res = subprocess.run([NODE, "-e", script], capture_output=True, text=True, check=True)
    out = json.loads(res.stdout)
    assert out["start"] is True
    assert out["afterStart"] == [[3, 5]]
    assert out["countKept"] == 3
    assert out["end"] is True
    assert out["afterEnd"] == [[3, 4]]
    assert out["countMoved"] == 3
    assert out["implicit"] == [[0, 3]]
    assert out["implicitMoved"] == [[1, 3]]
    assert out["none"] == [None, None]
    assert out["two"] == [[0, 1], [4, 5]]
    assert out["segments"] == [[0, 1], [2, 3]]
