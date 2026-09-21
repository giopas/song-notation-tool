"""
userpaths.py — where the user's own files live, and how that's remembered.

Songs used to be written to ./songs, i.e. inside the source tree. They were
gitignored, so they never got committed, but "not in the repo's history" is
not the same as "safe": a `git clean -xdf`, a fresh clone, or deleting the
checkout takes the whole folder with it. A person's songs are their data,
not part of this program, so they belong under their own home directory and
get whatever backup regime the rest of it has.

This module owns two things and nothing else:

  * the songs folder    — default ~/Documents/Song Notation Tool
  * the last-used export folder — so the Save panel opens where you were

Both are kept in a small JSON config outside the source tree too. Pure
stdlib, no UI import: the web server, the CLI, and the Tkinter app all read
the same answer.
"""

from __future__ import annotations

import json
import os
import shutil
import sys

APP_DIR_NAME = "Song Notation Tool"
CONFIG_NAME = "config.json"


# ==============================================================================
#  Locations
# ==============================================================================

def config_dir() -> str:
    """Per-user config directory, following each platform's convention."""
    home = os.path.expanduser("~")
    if sys.platform == "darwin":
        return os.path.join(home, "Library", "Application Support", APP_DIR_NAME)
    if os.name == "nt":
        base = os.environ.get("APPDATA") or os.path.join(home, "AppData", "Roaming")
        return os.path.join(base, APP_DIR_NAME)
    base = os.environ.get("XDG_CONFIG_HOME") or os.path.join(home, ".config")
    return os.path.join(base, "song-notation-tool")


def config_path() -> str:
    return os.path.join(config_dir(), CONFIG_NAME)


def default_songs_dir() -> str:
    return os.path.join(os.path.expanduser("~"), "Documents", APP_DIR_NAME)


def default_export_dir() -> str:
    return os.path.join(os.path.expanduser("~"), "Documents")


# ==============================================================================
#  Config read / write  (best-effort: a broken or unwritable config file
#  must never stop the app from opening — it just falls back to defaults)
# ==============================================================================

def load_config() -> dict:
    try:
        with open(config_path(), "r", encoding="utf-8") as f:
            data = json.load(f)
        return data if isinstance(data, dict) else {}
    except (OSError, ValueError):
        return {}


def save_config(cfg: dict) -> bool:
    try:
        os.makedirs(config_dir(), exist_ok=True)
        with open(config_path(), "w", encoding="utf-8") as f:
            json.dump(cfg, f, indent=2)
        return True
    except OSError:
        return False


def _set(key: str, value) -> bool:
    cfg = load_config()
    cfg[key] = value
    return save_config(cfg)


def songs_dir() -> str:
    """The configured songs folder, or the default."""
    d = load_config().get("songs_dir")
    return os.path.expanduser(d) if isinstance(d, str) and d.strip() else default_songs_dir()


def set_songs_dir(path: str) -> bool:
    return _set("songs_dir", os.path.abspath(os.path.expanduser(path)))


def last_export_dir() -> str:
    """Where the export Save panel should open. Falls back to ~/Documents
    once the remembered folder stops existing (an unplugged drive, say)."""
    d = load_config().get("last_export_dir")
    if isinstance(d, str) and os.path.isdir(os.path.expanduser(d)):
        return os.path.expanduser(d)
    return default_export_dir()


def set_last_export_dir(path: str) -> bool:
    return _set("last_export_dir", os.path.abspath(os.path.expanduser(path)))


# ==============================================================================
#  One-time migration out of the source tree
# ==============================================================================

def _sng_files(directory: str):
    try:
        return sorted(f for f in os.listdir(directory) if f.lower().endswith(".sng"))
    except OSError:
        return []


def migrate_legacy_songs(legacy_dir: str, target_dir: str):
    """
    Move any .sng files still sitting in the old in-repo folder to
    `target_dir`, and return the list of names moved.

    Deliberately conservative: it never overwrites a file already in the
    target (a name clash is skipped and reported), and it never deletes
    the old folder — the originals are moved, and anything that couldn't
    be moved is left exactly where it was.
    """
    legacy_dir = os.path.abspath(os.path.expanduser(legacy_dir))
    target_dir = os.path.abspath(os.path.expanduser(target_dir))
    if legacy_dir == target_dir or not os.path.isdir(legacy_dir):
        return [], []

    names = _sng_files(legacy_dir)
    if not names:
        return [], []

    try:
        os.makedirs(target_dir, exist_ok=True)
    except OSError:
        return [], names

    moved, skipped = [], []
    for name in names:
        src, dst = os.path.join(legacy_dir, name), os.path.join(target_dir, name)
        if os.path.exists(dst):
            skipped.append(name)
            continue
        try:
            shutil.move(src, dst)
            moved.append(name)
        except OSError:
            skipped.append(name)
    return moved, skipped


# ==============================================================================
#  Changing the songs folder — and taking the songs along
#
#  Pointing the app at a new folder used to leave every song behind in the
#  old one, so "change folder" looked like "my songs are gone". Now the
#  choice is explicit: move them, or just use the new folder as it is.
#
#  The move is the same conservative one the legacy migration uses: only
#  .sng files (the folder could be ~/Documents itself — nothing else in it
#  is ours to move), never overwrite a song already at the destination
#  (a clash is skipped and reported, and stays where it was), and never
#  delete the old folder.
# ==============================================================================

def song_files(directory: str) -> list:
    """The .sng files directly inside `directory`, sorted — the songs a
    move would take along."""
    return _sng_files(os.path.abspath(os.path.expanduser(directory)))


def plan_relocation(current_dir: str, new_dir: str) -> dict:
    """What moving from `current_dir` to `new_dir` would do, without doing
    it — for the question the app asks before it touches anything.

    {"current", "new", "same", "to_move": [...], "clashes": [...],
     "already_there": n, "writable": bool, "error": str}
    """
    cur = os.path.abspath(os.path.expanduser(current_dir))
    new = os.path.abspath(os.path.expanduser(new_dir))
    out = {"current": cur, "new": new, "same": cur == new, "to_move": [],
           "clashes": [], "already_there": 0, "writable": True, "error": ""}
    if os.path.exists(new) and not os.path.isdir(new):
        out.update(writable=False, error=f"{new} is a file, not a folder")
        return out
    parent = new if os.path.isdir(new) else os.path.dirname(new)
    while parent and not os.path.isdir(parent):
        parent = os.path.dirname(parent)
    if not parent or not os.access(parent, os.W_OK):
        out.update(writable=False, error=f"can't write to {new}")
        return out
    there = set(song_files(new)) if os.path.isdir(new) else set()
    out["already_there"] = len(there)
    if not out["same"]:
        here = song_files(cur)
        out["to_move"] = [n for n in here if n not in there]
        out["clashes"] = [n for n in here if n in there]
    return out


def relocate_songs(current_dir: str, new_dir: str, move: bool = True) -> dict:
    """Make `new_dir` the songs folder, moving the songs there if `move`.

    The config is only updated once the new folder exists and is usable,
    so a failed change leaves the app pointing where it was. Returns
    {"ok", "path", "moved": [...], "skipped": [...], "error"}.
    """
    plan = plan_relocation(current_dir, new_dir)
    new = plan["new"]
    if not plan["writable"]:
        return {"ok": False, "path": new, "moved": [], "skipped": [],
                "error": plan["error"]}
    try:
        os.makedirs(new, exist_ok=True)
    except OSError as exc:
        return {"ok": False, "path": new, "moved": [], "skipped": [],
                "error": str(exc)}
    moved, skipped = [], []
    if move and not plan["same"]:
        moved, skipped = migrate_legacy_songs(plan["current"], new)
    if not set_songs_dir(new):
        return {"ok": False, "path": new, "moved": moved, "skipped": skipped,
                "error": "couldn't save the setting"}
    return {"ok": True, "path": new, "moved": moved, "skipped": skipped,
            "error": ""}
