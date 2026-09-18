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


# ==============================================================================
#  Reference lookup — v0.20
# ==============================================================================

def test_find_section_matches_id_then_name():
    import model, songmap
    doc = model.new_document(title="T")
    sec = model.new_section("chorus1", "Interlude", "Interlude")
    doc["sections"] = [sec]
    assert songmap.find_section(doc, "chorus1") is sec
    assert songmap.find_section(doc, "Interlude") is sec
    assert songmap.find_section(doc, "interlude") is sec   # case-insensitive
    assert songmap.find_section(doc, "nope") is None


def test_section_display_name_falls_back_to_the_key():
    import model, songmap
    doc = model.new_document(title="T")
    doc["sections"] = [model.new_section("chorus1", "Interlude", "Interlude")]
    assert songmap.section_display_name(doc, "chorus1") == "Interlude"
    assert songmap.section_display_name(doc, "nope") == "nope"


def test_section_refs_to_finds_a_ref_written_by_name():
    import model, songmap
    doc = model.new_document(title="T")
    target = model.new_section("chorus1", "Interlude", "Interlude")
    user = model.new_section("s2", "Later", "Verse")
    user["items"] = [model.make_section_ref("Interlude")]
    doc["sections"] = [target, user]
    assert songmap.section_refs_to(doc, "chorus1") == ["Later"]


# ==============================================================================
#  User file locations — v0.20
# ==============================================================================

def test_songs_dir_defaults_under_the_user_home_not_the_repo(monkeypatch, tmp_path):
    import userpaths
    monkeypatch.setattr(userpaths, "config_path", lambda: str(tmp_path / "cfg.json"))
    d = userpaths.songs_dir()
    assert d == userpaths.default_songs_dir()
    assert "Documents" in d


def test_songs_dir_round_trips_through_the_config(monkeypatch, tmp_path):
    import userpaths
    monkeypatch.setattr(userpaths, "config_dir", lambda: str(tmp_path))
    monkeypatch.setattr(userpaths, "config_path", lambda: str(tmp_path / "cfg.json"))
    assert userpaths.set_songs_dir(str(tmp_path / "Tunes"))
    assert userpaths.songs_dir() == str(tmp_path / "Tunes")


def test_last_export_dir_falls_back_when_the_folder_is_gone(monkeypatch, tmp_path):
    import userpaths
    monkeypatch.setattr(userpaths, "config_dir", lambda: str(tmp_path))
    monkeypatch.setattr(userpaths, "config_path", lambda: str(tmp_path / "cfg.json"))
    gone = tmp_path / "unplugged-drive"
    gone.mkdir()
    userpaths.set_last_export_dir(str(gone))
    assert userpaths.last_export_dir() == str(gone)
    gone.rmdir()
    assert userpaths.last_export_dir() == userpaths.default_export_dir()


def test_broken_config_falls_back_to_defaults(monkeypatch, tmp_path):
    import userpaths
    bad = tmp_path / "cfg.json"
    bad.write_text("{not json at all")
    monkeypatch.setattr(userpaths, "config_path", lambda: str(bad))
    assert userpaths.load_config() == {}
    assert userpaths.songs_dir() == userpaths.default_songs_dir()


def test_migration_moves_songs_out_of_the_repo(tmp_path):
    import userpaths
    legacy = tmp_path / "repo" / "songs"; legacy.mkdir(parents=True)
    target = tmp_path / "Documents" / "Song Notation Tool"
    (legacy / "a.sng").write_text("{}")
    (legacy / "b.sng").write_text("{}")
    (legacy / "notes.txt").write_text("not a song")

    moved, skipped = userpaths.migrate_legacy_songs(str(legacy), str(target))
    assert sorted(moved) == ["a.sng", "b.sng"] and skipped == []
    assert (target / "a.sng").exists()
    assert not (legacy / "a.sng").exists()
    assert (legacy / "notes.txt").exists()      # untouched


def test_migration_never_overwrites_and_never_deletes(tmp_path):
    import userpaths
    legacy = tmp_path / "repo" / "songs"; legacy.mkdir(parents=True)
    target = tmp_path / "songs-target"; target.mkdir()
    (legacy / "clash.sng").write_text("old")
    (target / "clash.sng").write_text("newer")

    moved, skipped = userpaths.migrate_legacy_songs(str(legacy), str(target))
    assert moved == [] and skipped == ["clash.sng"]
    assert (target / "clash.sng").read_text() == "newer"   # not clobbered
    assert (legacy / "clash.sng").read_text() == "old"     # not deleted


def test_migration_is_a_no_op_when_source_and_target_match(tmp_path):
    import userpaths
    d = tmp_path / "songs"; d.mkdir(); (d / "a.sng").write_text("{}")
    assert userpaths.migrate_legacy_songs(str(d), str(d)) == ([], [])
    assert (d / "a.sng").exists()


# ==============================================================================
#  Reference display vs. storage — v0.20
# ==============================================================================

def _doc_renamed_section():
    """A section created as "Chorus" (id chorus1) and later renamed."""
    import model
    doc = model.new_document(title="T")
    doc["sections"] = [model.new_section("chorus1", "Interlude", "Interlude")]
    return doc


def test_display_ref_key_shows_the_current_name_not_the_minted_id():
    import songmap
    doc = _doc_renamed_section()
    assert songmap.display_ref_key(doc, "chorus1") == "Interlude"


def test_display_ref_key_underscores_a_name_with_spaces():
    import model, songmap
    doc = model.new_document(title="T")
    doc["sections"] = [model.new_section("s1", "Verse 2", "Verse")]
    key = songmap.display_ref_key(doc, "s1")
    assert key == "Verse_2"
    assert songmap.find_section(doc, key)["name"] == "Verse 2"   # round-trips


def test_display_ref_key_falls_back_to_the_id_when_the_name_cannot_be_written():
    import model, songmap
    doc = model.new_document(title="T")
    doc["sections"] = [model.new_section("s1", "4/4 breakdown!", "Custom")]
    assert songmap.display_ref_key(doc, "s1") == "s1"


def test_display_ref_key_falls_back_to_the_id_when_two_sections_share_a_name():
    import model, songmap
    doc = model.new_document(title="T")
    doc["sections"] = [model.new_section("a", "Verse", "Verse"),
                       model.new_section("b", "Verse", "Verse")]
    assert songmap.display_ref_key(doc, "a") == "a"


def test_display_ref_key_leaves_an_unresolvable_key_alone():
    import songmap
    assert songmap.display_ref_key(_doc_renamed_section(), "nope") == "nope"


def test_canonicalise_refs_stores_ids_whatever_was_typed():
    import model, songmap
    doc = _doc_renamed_section()
    for typed in ("Interlude", "interlude", "chorus1"):
        out = songmap.canonicalise_refs([model.make_section_ref(typed)], doc)
        assert out[0]["section"] == "chorus1"


def test_canonicalise_refs_leaves_an_unknown_target_as_typed():
    import model, songmap
    out = songmap.canonicalise_refs(
        [model.make_section_ref("not_yet")], _doc_renamed_section())
    assert out[0]["section"] == "not_yet"


def test_display_and_canonicalise_reach_into_groups():
    import model, songmap
    doc = _doc_renamed_section()
    items = [model.make_group([model.make_section_ref("chorus1")])]
    shown = songmap.display_items(items, doc)
    assert shown[0]["items"][0]["section"] == "Interlude"
    back = songmap.canonicalise_refs(shown, doc)
    assert back[0]["items"][0]["section"] == "chorus1"
    assert items[0]["items"][0]["section"] == "chorus1"    # input not mutated


def test_a_rename_changes_the_display_but_not_the_stored_reference():
    import model, songmap
    doc = _doc_renamed_section()
    ref = [model.make_section_ref("chorus1")]
    assert songmap.display_ref_key(doc, "chorus1") == "Interlude"
    doc["sections"][0]["name"] = "Bridge"
    assert songmap.display_ref_key(doc, "chorus1") == "Bridge"
    assert ref[0]["section"] == "chorus1"                  # still resolves
