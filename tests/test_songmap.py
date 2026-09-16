import os
import sys

import pytest

sys.path.insert(0, os.path.join(os.path.dirname(__file__), ".."))

import model
import songmap
from grammar import ParseError


def _doc_with_riff():
    doc = model.new_document(title="T", artist="A", key="B")
    doc["blocks"]["riff1"] = model.new_block("riff1", "Riff 1", bars=4)
    doc["blocks"]["riff2"] = model.new_block("riff2", "Riff 2", bars=4)

    verse = model.new_section("verse1", "Verse 1")
    verse["items"] = [model.make_block_ref("riff1", repeat=4)]
    doc["sections"].append(verse)

    chorus = model.new_section("chorus1", "Chorus 1")
    chorus["items"] = [model.make_group(
        [model.make_block_ref("riff1"), model.make_block_ref("riff2")])]
    doc["sections"].append(chorus)

    return doc


# ==============================================================================
#  block_usage / can_delete_block
# ==============================================================================

def test_block_usage_counts_each_section_once():
    doc = _doc_with_riff()
    usage = songmap.block_usage(doc)
    assert usage["riff1"] == ["Verse 1", "Chorus 1"]
    assert usage["riff2"] == ["Chorus 1"]


def test_block_usage_finds_refs_inside_groups():
    doc = _doc_with_riff()
    usage = songmap.block_usage(doc)
    assert "Chorus 1" in usage["riff2"]


def test_can_delete_block_blocked_when_referenced():
    doc = _doc_with_riff()
    ok, refs = songmap.can_delete_block(doc, "riff1")
    assert ok is False
    assert refs == ["Verse 1", "Chorus 1"]


def test_can_delete_block_allowed_when_unused():
    doc = _doc_with_riff()
    doc["blocks"]["riff3"] = model.new_block("riff3", "Riff 3")
    ok, refs = songmap.can_delete_block(doc, "riff3")
    assert ok is True
    assert refs == []


def test_unique_block_id_skips_existing():
    doc = _doc_with_riff()
    assert songmap.unique_block_id(doc) == "riff3"
    assert songmap.unique_block_id(doc, stem="x") == "x1"


# ==============================================================================
#  promote_to_riff
# ==============================================================================

def test_promote_to_riff_replaces_selection_with_block_id():
    line = "5A 5D [5A 5D]x2 riff1 x3"
    sel_start = line.index("[5A 5D]x2")
    sel_end = sel_start + len("[5A 5D]x2")
    new_line, riff_items = songmap.promote_to_riff(line, sel_start, sel_end, "riffnew")
    assert new_line == "5A 5D riffnew riff1 x3"
    assert len(riff_items) == 1
    assert riff_items[0]["kind"] == "group"


def test_promote_to_riff_inserts_boundary_space_when_missing():
    # No trailing space in the source before the next token — the old
    # off-by-whitespace bug fused "riffnew" directly onto "5D".
    line = "5A 5D [5A 5D 7A 9A]x25D 7A"
    sel_start = line.index("[5A 5D 7A 9A]x2")
    sel_end = sel_start + len("[5A 5D 7A 9A]x2")
    new_line, _ = songmap.promote_to_riff(line, sel_start, sel_end, "riffnew")
    assert new_line == "5A 5D riffnew 5D 7A"
    # Round-trips through the grammar as five separate items, not four —
    # the old bug fused "riffnew" and "5D" into one unparseable-as-two token.
    from grammar import parse_items
    items = parse_items(new_line)
    assert len(items) == 5
    assert items[2] == {"kind": "block_ref", "block": "riffnew", "repeat": 1, "transpose": 0}


def test_promote_to_riff_rejects_empty_selection():
    with pytest.raises(ValueError):
        songmap.promote_to_riff("5A 5D", 2, 2, "riffnew")


def test_promote_to_riff_propagates_parse_error():
    with pytest.raises(ParseError):
        songmap.promote_to_riff("5A [unclosed", 2, 12, "riffnew")


# ==============================================================================
#  move_section / reorder_sections
# ==============================================================================

def _three_section_doc():
    doc = model.new_document()
    for sid in ("a", "b", "c"):
        doc["sections"].append(model.new_section(sid, sid.upper()))
    return doc


def test_move_section_up_and_down():
    doc = _three_section_doc()
    new_idx = songmap.move_section(doc, 1, -1)
    assert new_idx == 0
    assert [s["id"] for s in doc["sections"]] == ["b", "a", "c"]

    new_idx = songmap.move_section(doc, 0, 1)
    assert new_idx == 1
    assert [s["id"] for s in doc["sections"]] == ["a", "b", "c"]


def test_move_section_clamps_at_boundary():
    doc = _three_section_doc()
    new_idx = songmap.move_section(doc, 0, -1)
    assert new_idx == 0
    assert [s["id"] for s in doc["sections"]] == ["a", "b", "c"]


def test_reorder_sections_to_arbitrary_position():
    doc = _three_section_doc()
    songmap.reorder_sections(doc, 0, 2)
    assert [s["id"] for s in doc["sections"]] == ["b", "c", "a"]


# ==============================================================================
#  duplicate_as_reference
# ==============================================================================

def test_duplicate_as_reference_inserts_section_ref_after_source():
    doc = _three_section_doc()
    new_sec = songmap.duplicate_as_reference(doc, 0, "a2", "A (repeat)")
    ids = [s["id"] for s in doc["sections"]]
    assert ids == ["a", "a2", "b", "c"]
    assert new_sec["items"] == [model.make_section_ref("a")]


# ==============================================================================
#  section_refs_to
# ==============================================================================

def test_section_refs_to_finds_referencing_sections():
    doc = _three_section_doc()
    doc["sections"][1]["items"] = [model.make_section_ref("a")]
    assert songmap.section_refs_to(doc, "a") == ["B"]
    assert songmap.section_refs_to(doc, "c") == []


# ==============================================================================
#  chart_items / measure_items split (render:"both" sections)
# ==============================================================================

def test_chart_and_measure_items_split_and_recombine():
    sec = model.new_section("bridge", "Bridge", render="both")
    sec["items"] = [
        model.make_token("A", 5),
        model.make_measure(8, {"E": "- - - - - - - -"}),
        model.make_token("B", 7),
    ]
    assert songmap.chart_items(sec) == [model.make_token("A", 5), model.make_token("B", 7)]
    assert songmap.measure_items(sec) == [model.make_measure(8, {"E": "- - - - - - - -"})]


def test_set_chart_items_preserves_existing_measures():
    sec = model.new_section("bridge", "Bridge", render="both")
    m = model.make_measure(8, {"E": "- 7 - - - - - -"})
    sec["items"] = [m]
    songmap.set_chart_items(sec, [model.make_token("A", 5)])
    assert sec["items"] == [model.make_token("A", 5), m]


def test_set_measure_items_preserves_existing_chart_items():
    sec = model.new_section("bridge", "Bridge", render="both")
    tok = model.make_token("A", 5)
    sec["items"] = [tok]
    songmap.set_measure_items(sec, [model.make_measure(8, {"E": "- - - - - - - -"})])
    assert sec["items"] == [tok, model.make_measure(8, {"E": "- - - - - - - -"})]
