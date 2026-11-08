"""A part chooses Chart, Tab, Both or Free, as a section does (v0.33.0).

A bass line under a guitar section is often easier to write as tab, and
some parts are only a note ("follow the kick"). These tests check that
each mode prints what it should in both exports, that an older part with
no mode is still a chart line, and that a part's tab cells are checked
like a section's.
"""
import os
import sys

sys.path.insert(0, os.path.join(os.path.dirname(__file__), ".."))

import export
import grammar
import model
import render


def _song(part):
    doc = model.new_document(title="Test Tune", artist="The Placeholders")
    sec = model.new_section("v1", "Verse", "Verse", instrument="Guitar (6-string)")
    sec["items"] = grammar.parse_items("G D Em C")
    sec["parts"] = [part]
    doc["sections"] = [sec]
    return doc, sec


def _measure(strings, beats=8):
    return {"kind": "measure", "beats": beats, "strings": strings}


def _texts(rows):
    return [r["text"] for r in rows]


def test_an_older_part_without_a_mode_is_a_chart_line():
    part = {"id": "p1", "instrument": "Bass (4-string)",
            "items": grammar.parse_items("G D")}
    doc, sec = _song(part)
    rows = render.printed_part_rows(doc, sec, None, 0)
    assert rows[0]["text"] == "Bass (4-string):"
    assert any("G" in t and "D" in t for t in _texts(rows[1:]))


def test_a_chart_part_is_stored_as_before():
    assert model.make_part("p1", "Bass (4-string)") == {
        "id": "p1", "instrument": "Bass (4-string)", "items": []}


def test_a_tab_part_prints_its_grid_and_not_its_chart_line():
    part = model.make_part("p1", "Bass (4-string)",
                           grammar.parse_items("A D") + [_measure({"E": "5 - 7 - 5 - 3 -"})],
                           render="tab")
    doc, sec = _song(part)
    text = "\n".join(_texts(render.printed_part_rows(doc, sec, None, 0)))
    assert "M1" in text and "E|" in text and "7" in text
    assert "A" not in text.replace("Bass", "")      # the chart line stays out


def test_a_both_part_prints_the_chart_then_the_grid():
    part = model.make_part("p1", "Bass (4-string)",
                           grammar.parse_items("A D") + [_measure({"E": "5 - - - - - - -"})],
                           render="both")
    doc, sec = _song(part)
    rows = render.printed_part_rows(doc, sec, None, 0)
    roles = [r["role"] for r in rows]
    assert render.ROLE_GRID in roles
    first_grid = roles.index(render.ROLE_GRID)
    assert any("A" in t for t in _texts(rows[1:first_grid]))


def test_a_free_part_prints_its_text_as_typed():
    part = model.make_part("p1", "Bass (4-string)", render="free",
                           free_text="Follow the kick\n  root notes only")
    doc, sec = _song(part)
    rows = render.printed_part_rows(doc, sec, None, 0)
    assert _texts(rows) == ["Bass (4-string):", "Follow the kick", "  root notes only"]


def test_an_empty_part_prints_nothing_in_any_mode():
    for mode in ("chart", "tab", "both", "free"):
        doc, sec = _song(model.make_part("p1", "Bass (4-string)", render=mode))
        assert render.printed_part_rows(doc, sec, None, 0) == []
        assert render.section_parts(sec) == []


def test_both_exports_carry_a_tab_part():
    part = model.make_part("p1", "Bass (4-string)",
                           [_measure({"E": "5 - 7 - 5 - 3 -"})], render="tab")
    doc, _sec = _song(part)
    txt, _warnings = export.prepare_export(doc, "txt")
    body = txt.decode("utf-8") if isinstance(txt, bytes) else txt
    assert "Bass (4-string):" in body and "E|" in body and "M1" in body
    pdf, _warnings = export.prepare_export(doc, "pdf")
    assert pdf[:5] == b"%PDF-"


def test_a_parts_tab_cells_are_checked_like_a_sections():
    part = model.make_part("p1", "Bass (4-string)",
                           [_measure({"E": "5 zz - - - - - -"})], render="tab")
    doc, _sec = _song(part)
    res = export.check_document(doc)
    assert any("Verse / Bass (4-string)" in w and "zz" in w for w in res["warnings"])
