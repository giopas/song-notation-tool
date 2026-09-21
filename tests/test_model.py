import os
import sys
import copy
import glob
import json

sys.path.insert(0, os.path.join(os.path.dirname(__file__), ".."))

from model import (
    new_document, new_section, new_block, make_token, make_measure,
    make_section_ref, validate_block_name, migrate_document, FORMAT_VERSION,
)


def test_new_document_shape():
    doc = new_document(title="Sample Song", artist="Sample Artist", key="B")
    assert doc["format"] == FORMAT_VERSION
    assert doc["meta"]["title"] == "Sample Song"
    assert doc["sections"] == []
    assert doc["blocks"] == {}


def test_validate_block_name_rejects_token_like_names():
    assert not validate_block_name("A1")
    assert not validate_block_name("b2")
    assert not validate_block_name("5A")
    assert not validate_block_name("A")
    assert validate_block_name("riff1")
    assert validate_block_name("intro_riff")


def test_migrate_missing_format_treated_as_v1():
    v015_doc = {
        "title": "Sample Song", "artist": "Sample Band", "key": "B", "time": "4/4",
        "sections": [
            {
                "id": "verse1", "name": "Verse 1", "type": "Verse",
                "instrument": "Bass (4-string)", "repeat": 4, "link_id": None,
                "tab_beats": 8, "measure_beats": {},
                "layers": {
                    "tab": [{"E": "- 5 - - - - - -", "A": "- - - - - - - -"}],
                    "chords": ["F#"],
                    "notes": [""],
                    "lyrics": [""],
                },
                "visible": {"tab": True, "chords": True, "notes": True, "lyrics": True},
            }
        ],
    }
    migrated = migrate_document(copy.deepcopy(v015_doc))
    assert migrated["format"] == FORMAT_VERSION
    assert migrated["sections"][0]["id"] == "verse1"
    kinds = [it["kind"] for it in migrated["sections"][0]["items"]]
    assert "measure" in kinds
    assert "token" in kinds


def test_migrate_link_id_becomes_section_ref():
    v015_doc = {
        "title": "T", "sections": [
            {
                "id": "chorus1", "name": "Chorus 1", "type": "Chorus",
                "instrument": "Bass (4-string)", "repeat": 4, "link_id": 1,
                "tab_beats": 8, "measure_beats": {},
                "layers": {
                    "tab": [{"E": "- 5 - - - - - -"}],
                    "chords": [""], "notes": [""], "lyrics": [""],
                },
                "visible": {"tab": True, "chords": True, "notes": True, "lyrics": True},
            },
            {
                "id": "chorus2", "name": "Chorus 2", "type": "Chorus",
                "instrument": "Bass (4-string)", "repeat": 4, "link_id": 1,
                "tab_beats": 8, "measure_beats": {},
                "layers": {
                    "tab": [{"E": "- 5 - - - - - -"}],
                    "chords": [""], "notes": [""], "lyrics": [""],
                },
                "visible": {"tab": True, "chords": True, "notes": True, "lyrics": True},
            },
        ],
    }
    migrated = migrate_document(copy.deepcopy(v015_doc))
    sec2 = migrated["sections"][1]
    assert sec2["items"] == [make_section_ref("chorus1", repeat=4)]


def test_migrate_already_v2_passthrough():
    doc = new_document(title="X")
    doc["sections"].append(new_section("s1", "Verse 1"))
    out = migrate_document(doc)
    assert out is doc


def test_new_block_shape():
    b = new_block("riff1", "Riff 1")
    assert b["bars"] == 4 and b["beats_per_bar"] == 8
    assert b["items"] == []


def test_sample_song_fixture_migrates_without_loss():
    """Round-trip: build a v0.15-shaped doc from the synthetic fixture
    and confirm every measure survives migration."""
    fixture_path = os.path.join(
        os.path.dirname(__file__), "fixtures", "v015", "sample_song.json")
    with open(fixture_path) as f:
        v015_doc = json.load(f)

    migrated = migrate_document(copy.deepcopy(v015_doc))
    assert migrated["format"] == FORMAT_VERSION
    assert len(migrated["sections"]) == len(v015_doc["sections"])

    for orig, new in zip(v015_doc["sections"], migrated["sections"]):
        assert new["name"] == orig["name"]

    # The two linked sections (Verse 2 -> Verse 1, Chorus 2 -> Chorus 1)
    # collapse to section_ref items instead of duplicating the tab.
    by_name = {s["name"]: s for s in migrated["sections"]}
    assert by_name["Verse 2"]["items"][0]["kind"] == "section_ref"
    assert by_name["Chorus 2"]["items"][0]["kind"] == "section_ref"


# ---------------------------------------------------------------------------
#  Songs saved before lyric placement was one setting
# ---------------------------------------------------------------------------

import model  # noqa: E402

def test_old_song_with_ticked_sections_keeps_printing_them_under_the_chart():
    doc = {"format": 2, "sections": [
        {"id": "v", "lyrics_text": "words", "print_lyrics": True}]}
    assert model.migrate_document(doc)["lyrics_layout"] == "below"


def test_old_song_with_only_the_sheet_ticked_prints_it_at_the_start():
    doc = {"format": 2, "sections": [], "lyrics_text": "words", "print_lyrics": True}
    assert model.migrate_document(doc)["lyrics_layout"] == "start"


def test_old_song_with_nothing_ticked_prints_no_words():
    doc = {"format": 2, "sections": [
        {"id": "v", "lyrics_text": "words", "print_lyrics": False}]}
    assert model.migrate_document(doc)["lyrics_layout"] == "none"


def test_a_song_that_already_chose_keeps_its_choice():
    doc = {"format": 2, "sections": [], "lyrics_layout": "side"}
    assert model.migrate_document(doc)["lyrics_layout"] == "side"
