"""
library.py — what the songs folder knows beyond the songs themselves.

Three things live here, none of them inside a song file:

  * hidden songs     — kept, but left out of the song list until unhidden
  * collections      — named, ordered lists of songs: a setlist, an album,
                       a "learn next" pile. A song can be in any number.
  * the trash        — a deleted song is moved to the "Deleted songs"
                       folder beside the others, and can be restored until
                       the trash is emptied.

They are kept in `library.json` in the songs folder, so they move with the
folder and a song file never changes because it joined a setlist. The
song files stay the truth: a name in the library whose file has gone is
skipped when the library is read, never an error.

Pure stdlib, no UI import: the web server and its tests use it directly.
"""

from __future__ import annotations

import datetime
import json
import os
import shutil
import threading
import uuid

LIBRARY_FILE = "library.json"
TRASH_DIR = "Deleted songs"
MAX_NAME = 80

_LOCK = threading.RLock()


class LibraryError(ValueError):
    """Something the user asked for that can't be done, in a sentence."""


def _now() -> str:
    return datetime.datetime.now().replace(microsecond=0).isoformat()


def _empty() -> dict:
    return {"version": 1, "hidden": [], "collections": [], "trash": {}}


def _free_name(folder: str, name: str) -> str:
    """`name`, or `name (2)`, `(3)`… — the first one not taken in `folder`."""
    stem, ext = os.path.splitext(name)
    candidate, n = name, 1
    while os.path.exists(os.path.join(folder, candidate)):
        n += 1
        candidate = f"{stem} ({n}){ext}"
    return candidate


def _clean_name(name) -> str:
    name = " ".join(str(name or "").split())
    if not name:
        raise LibraryError("A collection needs a name.")
    return name[:MAX_NAME]


class Library:
    """The library of one songs folder. Every call reads the file, changes
    it and writes it back, so two windows never work from stale copies."""

    def __init__(self, directory: str):
        self.dir = directory

    # ── the file ──────────────────────────────────────────────────────
    @property
    def path(self) -> str:
        return os.path.join(self.dir, LIBRARY_FILE)

    @property
    def trash_dir(self) -> str:
        return os.path.join(self.dir, TRASH_DIR)

    def load(self) -> dict:
        data = _empty()
        try:
            with open(self.path, "r", encoding="utf-8") as f:
                raw = json.load(f)
        except FileNotFoundError:
            return data
        except (OSError, ValueError):
            # Unreadable: keep it aside under a free name rather than
            # writing over it on the next change.
            self._set_aside_broken()
            return data
        if not isinstance(raw, dict):
            self._set_aside_broken()
            return data
        hidden = raw.get("hidden")
        data["hidden"] = [h for h in hidden if isinstance(h, str)] if isinstance(hidden, list) else []
        cols = raw.get("collections")
        for c in cols if isinstance(cols, list) else []:
            if not isinstance(c, dict) or not c.get("id"):
                continue
            songs = c.get("songs") if isinstance(c.get("songs"), list) else []
            data["collections"].append({
                "id": str(c["id"]),
                "name": str(c.get("name") or "Untitled"),
                "songs": [s for s in songs if isinstance(s, str)],
            })
        trash = raw.get("trash")
        if isinstance(trash, dict):
            data["trash"] = {k: v for k, v in trash.items() if isinstance(v, dict)}
        return data

    def _set_aside_broken(self):
        try:
            target = _free_name(self.dir, LIBRARY_FILE + ".unreadable")
            os.replace(self.path, os.path.join(self.dir, target))
        except OSError:
            pass

    def save(self, data: dict):
        os.makedirs(self.dir, exist_ok=True)
        tmp = self.path + ".saving"
        with open(tmp, "w", encoding="utf-8") as f:
            json.dump(data, f, indent=2, ensure_ascii=False)
        os.replace(tmp, self.path)

    def _exists(self, filename: str) -> bool:
        return os.path.isfile(os.path.join(self.dir, filename))

    def _collection(self, data: dict, cid: str) -> dict:
        for c in data["collections"]:
            if c["id"] == cid:
                return c
        raise LibraryError("That collection no longer exists.")

    # ── reading ───────────────────────────────────────────────────────
    def view(self) -> dict:
        """The library as the app shows it: collections with only the songs
        that still exist, in order, the hidden songs, and the trash size."""
        with _LOCK:
            data = self.load()
        return {
            "collections": [
                {"id": c["id"], "name": c["name"],
                 "songs": [s for s in c["songs"] if self._exists(s)]}
                for c in data["collections"]],
            "hidden": [h for h in data["hidden"] if self._exists(h)],
            "trash_count": len(self.trash_files()),
        }

    def song_flags(self) -> dict:
        """{filename: {"hidden": bool, "collections": [ids]}} for every
        song the library mentions."""
        with _LOCK:
            data = self.load()
        out = {}
        for h in data["hidden"]:
            out.setdefault(h, {"hidden": False, "collections": []})["hidden"] = True
        for c in data["collections"]:
            for s in c["songs"]:
                out.setdefault(s, {"hidden": False, "collections": []})["collections"].append(c["id"])
        return out

    def collection(self, cid: str) -> dict:
        for c in self.view()["collections"]:
            if c["id"] == cid:
                return c
        raise LibraryError("That collection no longer exists.")

    # ── hiding ────────────────────────────────────────────────────────
    def set_hidden(self, filename: str, hidden: bool = True):
        with _LOCK:
            data = self.load()
            data["hidden"] = [h for h in data["hidden"] if h != filename]
            if hidden:
                if not self._exists(filename):
                    raise LibraryError("That song no longer exists.")
                data["hidden"].append(filename)
            self.save(data)

    # ── collections ───────────────────────────────────────────────────
    def create_collection(self, name: str, songs=None) -> dict:
        with _LOCK:
            data = self.load()
            col = {"id": "c" + uuid.uuid4().hex[:10], "name": _clean_name(name),
                   "songs": []}
            for s in songs or []:
                if s not in col["songs"] and self._exists(s):
                    col["songs"].append(s)
            data["collections"].append(col)
            self.save(data)
            return dict(col)

    def rename_collection(self, cid: str, name: str):
        with _LOCK:
            data = self.load()
            self._collection(data, cid)["name"] = _clean_name(name)
            self.save(data)

    def delete_collection(self, cid: str):
        """Removes the list only. Its songs stay where they are."""
        with _LOCK:
            data = self.load()
            self._collection(data, cid)
            data["collections"] = [c for c in data["collections"] if c["id"] != cid]
            self.save(data)

    def add_songs(self, cid: str, filenames):
        """Add songs at the end of a collection; one already in it stays put."""
        with _LOCK:
            data = self.load()
            col = self._collection(data, cid)
            for s in filenames:
                if s not in col["songs"] and self._exists(s):
                    col["songs"].append(s)
            self.save(data)

    def remove_song(self, cid: str, filename: str):
        with _LOCK:
            data = self.load()
            col = self._collection(data, cid)
            col["songs"] = [s for s in col["songs"] if s != filename]
            self.save(data)

    def move_song(self, cid: str, filename: str, to_index: int):
        """Move a song to position `to_index` (0 = first) among the songs
        of the collection that still exist."""
        with _LOCK:
            data = self.load()
            col = self._collection(data, cid)
            present = [s for s in col["songs"] if self._exists(s)]
            if filename not in present:
                raise LibraryError("That song is not in this collection.")
            present.remove(filename)
            to_index = max(0, min(len(present), int(to_index)))
            present.insert(to_index, filename)
            col["songs"] = present
            self.save(data)

    def set_order(self, cid: str, filenames):
        """Put the collection's songs in the order given. Songs left out of
        `filenames` keep their place at the end, so a stale list can never
        drop a song from a setlist."""
        with _LOCK:
            data = self.load()
            col = self._collection(data, cid)
            present = [s for s in col["songs"] if self._exists(s)]
            ordered = [s for s in dict.fromkeys(filenames) if s in present]
            col["songs"] = ordered + [s for s in present if s not in ordered]
            self.save(data)

    # ── following a song file ─────────────────────────────────────────
    def rename_song(self, old: str, new: str):
        """A song was saved under a new file name: follow it."""
        if old == new:
            return
        with _LOCK:
            data = self.load()
            changed = False
            if old in data["hidden"]:
                data["hidden"] = [new if h == old else h for h in data["hidden"]]
                changed = True
            for c in data["collections"]:
                if old in c["songs"]:
                    c["songs"] = [new if s == old else s for s in c["songs"]]
                    changed = True
            if changed:
                self.save(data)

    # ── the trash ─────────────────────────────────────────────────────
    def trash_files(self) -> list:
        try:
            return sorted(f for f in os.listdir(self.trash_dir)
                          if f.lower().endswith(".sng"))
        except OSError:
            return []

    def trash_song(self, filename: str, title: str = "", artist: str = "") -> str:
        """Move a song to the trash. Returns its name there. What it belonged
        to (hidden, which collections and where) is remembered, so Restore
        puts it back the way it was."""
        src = os.path.join(self.dir, filename)
        if not os.path.isfile(src):
            raise LibraryError("That song no longer exists.")
        with _LOCK:
            os.makedirs(self.trash_dir, exist_ok=True)
            name = _free_name(self.trash_dir, filename)
            os.replace(src, os.path.join(self.trash_dir, name))
            data = self.load()
            record = {"original": filename, "deleted_at": _now(),
                      "title": title, "artist": artist,
                      "hidden": filename in data["hidden"], "collections": {}}
            data["hidden"] = [h for h in data["hidden"] if h != filename]
            for c in data["collections"]:
                if filename in c["songs"]:
                    record["collections"][c["id"]] = c["songs"].index(filename)
                    c["songs"] = [s for s in c["songs"] if s != filename]
            data["trash"][name] = record
            self.save(data)
            return name

    def trash_view(self) -> list:
        """The songs in the trash, the most recently deleted first."""
        with _LOCK:
            data = self.load()
        out = []
        for name in self.trash_files():
            rec = data["trash"].get(name, {})
            out.append({
                "name": name,
                "original": rec.get("original") or name,
                "title": rec.get("title") or "",
                "artist": rec.get("artist") or "",
                "deleted_at": rec.get("deleted_at") or "",
            })
        out.sort(key=lambda r: (r["deleted_at"], r["name"]), reverse=True)
        return out

    def _trash_path(self, name: str) -> str:
        if not name or os.path.basename(name) != name or name.startswith("."):
            raise LibraryError("No such song in the trash.")
        path = os.path.join(self.trash_dir, name)
        if not os.path.isfile(path):
            raise LibraryError("No such song in the trash.")
        return path

    def restore(self, name: str) -> str:
        """Put a song back. It takes its old name, or the next free one if
        another song has that name now. Returns the name it got."""
        with _LOCK:
            path = self._trash_path(name)
            data = self.load()
            rec = data["trash"].pop(name, {})
            target = _free_name(self.dir, rec.get("original") or name)
            os.replace(path, os.path.join(self.dir, target))
            if rec.get("hidden"):
                data["hidden"].append(target)
            for c in data["collections"]:
                idx = (rec.get("collections") or {}).get(c["id"])
                if idx is not None and target not in c["songs"]:
                    c["songs"].insert(max(0, min(len(c["songs"]), int(idx))), target)
            self.save(data)
            return target

    def purge(self, name: str):
        """Delete one song from the trash for good."""
        with _LOCK:
            path = self._trash_path(name)
            os.remove(path)
            data = self.load()
            data["trash"].pop(name, None)
            self.save(data)

    def empty_trash(self) -> int:
        """Delete every song in the trash for good. Returns how many."""
        with _LOCK:
            names = self.trash_files()
            for n in names:
                try:
                    os.remove(os.path.join(self.trash_dir, n))
                except OSError:
                    pass
            data = self.load()
            data["trash"] = {}
            self.save(data)
            return len(names)


def carry_over(current_dir: str, new_dir: str) -> dict:
    """Take the library and the trash along when the songs folder moves.

    Same rules as the songs: nothing at the destination is overwritten.
    A library already there is kept, and ours stays behind; a song in our
    trash whose name is taken in the new trash stays where it was.
    Returns {"library": bool, "trash": [names moved], "skipped": [...]}.
    """
    out = {"library": False, "trash": [], "skipped": []}
    if os.path.abspath(current_dir) == os.path.abspath(new_dir):
        return out
    with _LOCK:
        src = os.path.join(current_dir, LIBRARY_FILE)
        dst = os.path.join(new_dir, LIBRARY_FILE)
        if os.path.isfile(src) and not os.path.exists(dst):
            shutil.move(src, dst)
            out["library"] = True
        old_trash = os.path.join(current_dir, TRASH_DIR)
        new_trash = os.path.join(new_dir, TRASH_DIR)
        for name in Library(current_dir).trash_files():
            if os.path.exists(os.path.join(new_trash, name)):
                out["skipped"].append(name)
                continue
            os.makedirs(new_trash, exist_ok=True)
            shutil.move(os.path.join(old_trash, name), os.path.join(new_trash, name))
            out["trash"].append(name)
    return out
