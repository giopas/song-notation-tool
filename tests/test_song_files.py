"""A song file is named after the song, and follows it (v0.26)."""
import json
import os
import sys

sys.path.insert(0, os.path.join(os.path.dirname(__file__), ".."))

import pytest

import model
from constants import song_file_name, song_file_stem
from examples import example_document
import webserver


@pytest.fixture
def store(tmp_path):
    return webserver.SongStore(str(tmp_path))


def _doc(title, artist=""):
    return model.new_document(title=title, artist=artist)


def _files(store):
    return sorted(f for f in os.listdir(store.dir) if not f.startswith("."))


# ── the naming rule ─────────────────────────────────────────────────────

def test_name_is_artist_dash_title_with_spaces():
    assert song_file_name(_doc("Nutshell", "Alice in Chains")) == "Alice in Chains - Nutshell.sng"
    assert song_file_name(_doc("Nutshell")) == "Nutshell.sng"
    assert song_file_stem(_doc("")) == ""


def test_brackets_apostrophes_and_accents_survive():
    """They used to be stripped — "(MTV Unplugged)" lost its brackets."""
    for name in ("Nirvana - The Man Who Sold The World (MTV Unplugged).sng",
                 "Guns N' Roses - Don't Cry.sng", "Più bella cosa.sng"):
        assert webserver._safe_filename(name) == name


def test_nothing_can_step_outside_the_folder():
    assert webserver._safe_filename("../../etc/passwd") == "passwd.sng"
    assert webserver._safe_filename(".hidden") == "hidden.sng"
    assert "/" not in webserver._safe_filename("AC/DC - Thunderstruck")


# ── renaming on save ────────────────────────────────────────────────────

def test_the_example_turned_into_nutshell_is_renamed_on_save(store):
    """Exactly what happened: the file kept "Example Song.sng" after the
    song inside it became Nutshell."""
    store.save("Example Song.sng", example_document())
    doc = store.load("Example Song.sng")
    doc["meta"].update(title="Nutshell (MTV Unplugged)", artist="Alice in Chains")
    saved = store.save_song("Example Song.sng", doc)
    assert saved == "Alice in Chains - Nutshell (MTV Unplugged).sng"
    assert _files(store) == [saved]                  # old name gone, no copy
    assert store.load(saved)["meta"]["title"] == "Nutshell (MTV Unplugged)"


def test_saving_under_the_same_name_does_not_rename(store):
    doc = _doc("Nutshell", "Alice in Chains")
    store.save("Alice in Chains - Nutshell.sng", doc)
    assert store.save_song("Alice in Chains - Nutshell.sng", doc) == \
        "Alice in Chains - Nutshell.sng"
    assert _files(store) == ["Alice in Chains - Nutshell.sng"]


def test_another_song_is_never_overwritten(store):
    store.save("Nirvana - Lithium.sng", _doc("Lithium", "Nirvana"))
    other = _doc("Lithium", "Nirvana")
    other["sections"].append(model.new_section("x", "Mine"))
    store.save("draft.sng", other)
    saved = store.save_song("draft.sng", other)
    assert saved == "Nirvana - Lithium (2).sng"
    assert store.load("Nirvana - Lithium.sng")["sections"] == []   # untouched


def test_a_song_does_not_become_a_2_of_itself(store):
    doc = _doc("Lithium", "Nirvana")
    store.save("Nirvana - Lithium.sng", doc)
    for _ in range(3):
        assert store.save_song("Nirvana - Lithium.sng", doc) == "Nirvana - Lithium.sng"


def test_a_capitalisation_change_renames_the_file(store):
    store.save("nirvana - lithium.sng", _doc("lithium", "nirvana"))
    saved = store.save_song("nirvana - lithium.sng", _doc("Lithium", "Nirvana"))
    assert saved == "Nirvana - Lithium.sng"
    assert _files(store) == ["Nirvana - Lithium.sng"]


def test_a_song_with_no_title_keeps_its_name(store):
    store.save("scratch.sng", _doc(""))
    assert store.save_song("scratch.sng", _doc("")) == "scratch.sng"


def test_save_leaves_no_temporary_file(store):
    store.save("A.sng", _doc("A"))
    assert _files(store) == ["A.sng"]


# ── "Open example song" ─────────────────────────────────────────────────

def _open_example(store):
    """Drive the real route handler's logic through the store."""
    fresh = example_document()
    for entry in store.list_songs():
        if webserver._is_pristine_example(store.load(entry["filename"]), fresh):
            return entry["filename"]
    name = store.free_name(song_file_name(fresh))
    store.save(name, fresh)
    return name


def test_the_example_is_reused_while_untouched(store):
    a = _open_example(store)
    b = _open_example(store)
    assert a == b and len(_files(store)) == 1


def test_an_edited_example_is_never_handed_back_as_the_example(store):
    first = _open_example(store)
    doc = store.load(first)
    doc["meta"]["title"] = "Nutshell"
    store.save_song(first, doc)
    second = _open_example(store)
    assert store.load(second)["meta"]["title"] == example_document()["meta"]["title"]
    assert len(_files(store)) == 2


# ── pywebview's file-dialog constants ───────────────────────────────────

def test_dialog_kind_prefers_the_new_pywebview_name():
    """pywebview 5+ warns on webview.FOLDER_DIALOG; FileDialog.FOLDER is
    the current name. Old pywebview only has the old one."""
    import types
    new = types.SimpleNamespace(
        FileDialog=types.SimpleNamespace(FOLDER="new-folder", SAVE="new-save"),
        FOLDER_DIALOG="old-folder", SAVE_DIALOG="old-save")
    old = types.SimpleNamespace(FOLDER_DIALOG="old-folder", SAVE_DIALOG="old-save")
    assert webserver._dialog_kind(new, "FOLDER") == "new-folder"
    assert webserver._dialog_kind(new, "SAVE") == "new-save"
    assert webserver._dialog_kind(old, "FOLDER") == "old-folder"
    assert webserver._dialog_kind(old, "SAVE") == "old-save"
