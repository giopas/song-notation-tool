"""Changing the songs folder, and taking the songs along (v0.25)."""
import os
import sys

sys.path.insert(0, os.path.join(os.path.dirname(__file__), ".."))

import pytest

import userpaths


@pytest.fixture
def cfg(monkeypatch, tmp_path):
    """Keep the real config file out of it."""
    monkeypatch.setattr(userpaths, "config_dir", lambda: str(tmp_path / "cfg"))
    monkeypatch.setattr(userpaths, "config_path",
                        lambda: str(tmp_path / "cfg" / "config.json"))
    return tmp_path


def _songs(d, *names, extra=()):
    os.makedirs(d, exist_ok=True)
    for n in names:
        with open(os.path.join(d, n), "w") as f:
            f.write(f'{{"format": 2, "who": "{d}"}}')
    for n in extra:
        with open(os.path.join(d, n), "w") as f:
            f.write("not a song")


def test_plan_says_what_would_move_without_moving_it(cfg):
    old, new = cfg / "Old", cfg / "New"
    _songs(old, "A.sng", "B.sng")
    plan = userpaths.plan_relocation(str(old), str(new))
    assert plan["to_move"] == ["A.sng", "B.sng"]
    assert plan["writable"] and not plan["same"]
    assert sorted(os.listdir(old)) == ["A.sng", "B.sng"]   # untouched
    assert not new.exists()


def test_move_takes_the_songs_and_switches_the_folder(cfg):
    old, new = cfg / "Old", cfg / "iCloud" / "Songs"
    _songs(old, "A.sng", "B.sng")
    res = userpaths.relocate_songs(str(old), str(new), move=True)
    assert res["ok"] and res["moved"] == ["A.sng", "B.sng"]
    assert sorted(os.listdir(new)) == ["A.sng", "B.sng"]
    assert userpaths.songs_dir() == str(new)


def test_only_song_files_move(cfg):
    """The folder could be ~/Documents itself; nothing else is ours."""
    old, new = cfg / "Docs", cfg / "New"
    _songs(old, "A.sng", extra=("taxes.pdf", "notes.txt"))
    userpaths.relocate_songs(str(old), str(new), move=True)
    assert sorted(os.listdir(old)) == ["notes.txt", "taxes.pdf"]
    assert os.listdir(new) == ["A.sng"]


def test_a_song_already_there_is_never_overwritten(cfg):
    old, new = cfg / "Old", cfg / "New"
    _songs(old, "A.sng", "B.sng")
    _songs(new, "B.sng")
    plan = userpaths.plan_relocation(str(old), str(new))
    assert plan["clashes"] == ["B.sng"] and plan["to_move"] == ["A.sng"]
    res = userpaths.relocate_songs(str(old), str(new), move=True)
    assert res["moved"] == ["A.sng"] and res["skipped"] == ["B.sng"]
    with open(new / "B.sng") as f:
        assert str(new) in f.read()          # the destination's own copy
    assert os.listdir(old) == ["B.sng"]      # ours stays where it was


def test_just_use_this_folder_moves_nothing(cfg):
    old, new = cfg / "Old", cfg / "New"
    _songs(old, "A.sng")
    res = userpaths.relocate_songs(str(old), str(new), move=False)
    assert res["ok"] and res["moved"] == []
    assert os.listdir(old) == ["A.sng"]
    assert userpaths.songs_dir() == str(new)


def test_the_old_folder_is_never_deleted(cfg):
    old, new = cfg / "Old", cfg / "New"
    _songs(old, "A.sng")
    userpaths.relocate_songs(str(old), str(new), move=True)
    assert old.is_dir()


def test_a_file_is_refused_and_the_setting_left_alone(cfg):
    old = cfg / "Old"
    _songs(old, "A.sng")
    userpaths.set_songs_dir(str(old))
    not_a_folder = cfg / "file.txt"
    not_a_folder.write_text("x")
    res = userpaths.relocate_songs(str(old), str(not_a_folder), move=True)
    assert not res["ok"]
    assert userpaths.songs_dir() == str(old)
    assert os.listdir(old) == ["A.sng"]


def test_same_folder_is_a_no_op(cfg):
    old = cfg / "Old"
    _songs(old, "A.sng")
    plan = userpaths.plan_relocation(str(old), str(old))
    assert plan["same"] and plan["to_move"] == []
