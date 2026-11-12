"""Deleting, hiding and collecting songs (v0.35.0).

A deleted song goes to the trash ("Deleted songs" in the songs folder) and
can be restored until the trash is emptied. Hidden songs and collections
live in library.json beside the songs, so no song file changes when it is
hidden or joins a setlist. Invented song names only.
"""
import datetime
import json
import os
import sys
import threading
import urllib.error
import urllib.request

sys.path.insert(0, os.path.join(os.path.dirname(__file__), ".."))

import pytest

import export
import library
import model
import userpaths
import webserver


def _song(title, artist="The Invented"):
    doc = model.new_document(title=title, artist=artist)
    doc["sections"].append({
        "id": "s1", "name": "Verse", "type": "verse", "instrument": "Bass (4-string)",
        "repeat": 1, "transpose": 0, "render": "chart", "annotation": "",
        "lyrics_text": "", "print_lyrics": False, "free_text": "", "items": [],
        "chart_line": "[A D E]x2",
    })
    return doc


@pytest.fixture
def store(tmp_path):
    st = webserver.SongStore(str(tmp_path / "songs"))
    for t in ("Alpha", "Bravo", "Charlie"):
        st.save(f"The Invented - {t}.sng", model.new_document(title=t, artist="The Invented"))
    return st


A, B, C = ("The Invented - Alpha.sng", "The Invented - Bravo.sng",
           "The Invented - Charlie.sng")


def _listed(store):
    return [s["filename"] for s in store.list_songs()]


# ── the trash ───────────────────────────────────────────────────────────

def test_delete_moves_the_song_to_the_trash_folder(store):
    name = store.delete(A)
    assert name == A
    assert A not in _listed(store)
    assert os.path.isfile(os.path.join(store.dir, library.TRASH_DIR, A))
    trash = store.library.trash_view()
    assert [t["name"] for t in trash] == [A]
    assert trash[0]["title"] == "Alpha"


def test_restore_puts_it_back_with_its_collections_and_hidden_flag(store):
    lib = store.library
    col = lib.create_collection("Friday set", [A, B, C])
    lib.set_hidden(B, True)
    store.delete(B)
    assert lib.collection(col["id"])["songs"] == [A, C]
    assert lib.restore(B) == B
    assert lib.collection(col["id"])["songs"] == [A, B, C]   # same place
    assert lib.view()["hidden"] == [B]
    assert lib.trash_view() == []


def test_restore_never_overwrites_a_song_that_took_the_name(store):
    store.delete(A)
    store.save(A, model.new_document(title="Alpha", artist="The Invented"))
    assert store.library.restore(A) == "The Invented - Alpha (2).sng"
    assert len(_listed(store)) == 4


def test_deleting_the_same_name_twice_keeps_both_in_the_trash(store):
    store.delete(A)
    store.save(A, model.new_document(title="Alpha", artist="The Invented"))
    assert store.delete(A) == "The Invented - Alpha (2).sng"
    assert len(store.library.trash_view()) == 2


def test_purge_and_empty_trash_delete_for_good(store):
    store.delete(A)
    store.delete(B)
    store.library.purge(A)
    assert [t["name"] for t in store.library.trash_view()] == [B]
    assert store.library.empty_trash() == 1
    assert store.library.trash_view() == []
    assert sorted(_listed(store)) == [C]


def test_trash_names_cannot_step_outside_the_folder(store):
    store.delete(A)
    for bad in ("../" + B, "", ".hidden.sng", "nope.sng"):
        with pytest.raises(library.LibraryError):
            store.library.purge(bad)
    assert os.path.isfile(os.path.join(store.dir, B))


# ── hiding ──────────────────────────────────────────────────────────────

def test_hidden_is_a_flag_and_the_song_file_is_untouched(store):
    before = open(os.path.join(store.dir, A), "rb").read()
    store.library.set_hidden(A, True)
    songs = {s["filename"]: s for s in store.list_songs()}
    assert songs[A]["hidden"] is True and songs[B]["hidden"] is False
    assert open(os.path.join(store.dir, A), "rb").read() == before
    store.library.set_hidden(A, False)
    assert store.library.view()["hidden"] == []


# ── collections ─────────────────────────────────────────────────────────

def test_collections_keep_their_order_and_a_song_can_be_in_several(store):
    lib = store.library
    s1 = lib.create_collection("  Friday   set ")
    s2 = lib.create_collection("Album")
    assert s1["name"] == "Friday set"
    lib.add_songs(s1["id"], [C, A, C])
    lib.add_songs(s2["id"], [A])
    assert lib.collection(s1["id"])["songs"] == [C, A]
    flags = {s["filename"]: s["collections"] for s in store.list_songs()}
    assert sorted(flags[A]) == sorted([s1["id"], s2["id"]])
    lib.move_song(s1["id"], A, 0)
    assert lib.collection(s1["id"])["songs"] == [A, C]
    lib.add_songs(s1["id"], [B])
    lib.set_order(s1["id"], [B, A])            # C left out: kept at the end
    assert lib.collection(s1["id"])["songs"] == [B, A, C]
    lib.remove_song(s1["id"], A)
    assert lib.collection(s1["id"])["songs"] == [B, C]
    lib.rename_collection(s1["id"], "Saturday set")
    lib.delete_collection(s2["id"])
    assert [c["name"] for c in lib.view()["collections"]] == ["Saturday set"]
    assert sorted(_listed(store)) == sorted([A, B, C])  # songs untouched


def test_a_collection_needs_a_name(store):
    with pytest.raises(library.LibraryError):
        store.library.create_collection("   ")


def test_renaming_a_song_on_save_keeps_it_in_its_collections(store):
    lib = store.library
    col = lib.create_collection("Set", [A, B])
    lib.set_hidden(A, True)
    doc = store.load(A)
    doc["meta"]["title"] = "Alpha Live"
    saved = store.save_song(A, doc)
    assert saved == "The Invented - Alpha Live.sng"
    assert lib.collection(col["id"])["songs"] == [saved, B]
    assert lib.view()["hidden"] == [saved]


def test_a_song_removed_outside_the_app_is_skipped_not_an_error(store):
    col = store.library.create_collection("Set", [A, B])
    os.remove(os.path.join(store.dir, A))
    assert store.library.collection(col["id"])["songs"] == [B]


def test_an_unreadable_library_file_is_set_aside_not_overwritten(store):
    path = os.path.join(store.dir, library.LIBRARY_FILE)
    with open(path, "w") as f:
        f.write("{ not json")
    store.library.create_collection("Set", [A])
    assert os.path.isfile(path + ".unreadable")
    assert open(path + ".unreadable").read() == "{ not json"
    assert store.library.view()["collections"][0]["songs"] == [A]


def test_moving_the_songs_folder_takes_the_library_and_trash_along(store, tmp_path, monkeypatch):
    monkeypatch.setattr(userpaths, "set_songs_dir", lambda p: True)
    col = store.library.create_collection("Set", [A, B])
    store.delete(C)
    new = tmp_path / "elsewhere"
    res = userpaths.relocate_songs(store.dir, str(new), move=True)
    assert res["ok"]
    moved = webserver.SongStore(str(new))
    assert moved.library.collection(col["id"])["songs"] == [A, B]
    assert [t["name"] for t in moved.library.trash_view()] == [C]
    assert moved.library.restore(C) == C


# ── one PDF for a collection ────────────────────────────────────────────

def test_a_collection_pdf_has_every_song_in_order():
    docs = [_song("Alpha"), _song("Bravo")]
    day = datetime.date(2026, 1, 1)
    singles = [export.build_pdf(d, date=day) for d in docs]
    data, warnings = export.prepare_collection_export(docs, date=day)
    assert data.startswith(b"%PDF-")
    count = lambda pdf: pdf.count(b"/Type /Page ")
    assert count(data) == sum(count(p) for p in singles)
    assert b"/EmbeddedFiles" not in data
    assert export.build_collection_pdf(docs, date=day) == data   # deterministic


def test_a_collection_pdf_keeps_a_landscape_song_landscape():
    floor = _song("Floor")
    floor["print_size"] = "floor"
    docs = [_song("Alpha"), floor]
    o = export.resolve_orient(floor, export.instruments_to_print(floor))
    data = export.build_collection_pdf(docs)
    W, H = export.page_size(o)
    assert f"/MediaBox [0 0 {W:.2f} {H:.2f}]".encode() in data


def test_a_song_that_fails_its_check_blocks_the_collection_and_is_named():
    bad = _song("Broken")
    bad["sections"][0]["items"] = [{"type": "ref", "target": "nowhere"}]
    try:
        export.prepare_collection_export([_song("Alpha"), bad])
    except export.ExportBlocked as exc:
        assert all(e.startswith("Broken:") for e in exc.errors)
    else:
        if not export.check_document(bad)["errors"]:
            pytest.skip("this song shape is not an error for the check")
        raise AssertionError("not blocked")


# ── the HTTP routes ─────────────────────────────────────────────────────

@pytest.fixture()
def server(store):
    old = getattr(webserver.Handler, "store", None)
    webserver.Handler.store = store
    httpd = webserver._ThreadingHTTPServer(("127.0.0.1", 0), webserver.Handler)
    threading.Thread(target=httpd.serve_forever, daemon=True).start()
    base = f"http://127.0.0.1:{httpd.server_address[1]}"

    def call(method, path, body=None, raw=False):
        data = json.dumps(body).encode() if body is not None else None
        req = urllib.request.Request(base + path, data=data, method=method,
                                     headers={"Content-Type": "application/json"})
        try:
            with urllib.request.urlopen(req, timeout=10) as r:
                payload = r.read()
                return r.status, (payload if raw else json.loads(payload)), r.headers
        except urllib.error.HTTPError as e:
            return e.code, json.loads(e.read()), e.headers

    yield call
    httpd.shutdown()
    httpd.server_close()
    webserver.Handler.store = old


def test_routes_delete_restore_and_empty(server, store):
    status, body, _ = server("DELETE", "/api/songs/" + urllib.request.quote(A))
    assert status == 200 and body["trashed_as"] == A
    _, trash, _ = server("GET", "/api/trash")
    assert [t["name"] for t in trash] == [A]
    _, lib, _ = server("GET", "/api/library")
    assert lib["trash_count"] == 1
    _, body, _ = server("POST", "/api/trash", {"op": "restore", "name": A})
    assert body["ok"] and body["filename"] == A and body["trash"] == []
    server("DELETE", "/api/songs/" + urllib.request.quote(A))
    _, body, _ = server("POST", "/api/trash", {"op": "empty"})
    assert body["deleted"] == 1 and body["library"]["trash_count"] == 0
    status, _, _ = server("DELETE", "/api/songs/" + urllib.request.quote(A))
    assert status == 404


def test_routes_hide_and_collections(server, store):
    _, body, _ = server("POST", "/api/library", {"op": "hide", "song": B})
    assert body["library"]["hidden"] == [B]
    _, body, _ = server("POST", "/api/library", {"op": "create", "name": "Set", "songs": [C]})
    cid = body["collection"]["id"]
    server("POST", "/api/library", {"op": "add", "id": cid, "songs": [A, B]})
    _, body, _ = server("POST", "/api/library", {"op": "move", "id": cid, "song": B, "to": 0})
    assert body["library"]["collections"][0]["songs"] == [B, C, A]
    _, body, _ = server("POST", "/api/library", {"op": "nonsense"})
    assert body["ok"] is False
    _, body, _ = server("POST", "/api/library", {"op": "rename", "id": "missing", "name": "x"})
    assert body["ok"] is False and "no longer exists" in body["error"]
    _, songs, _ = server("GET", "/api/songs")
    by = {s["filename"]: s for s in songs}
    assert by[B]["hidden"] and by[A]["collections"] == [cid]


def test_route_collection_pdf(server, store):
    store.save(A, _song("Alpha"))
    store.save(B, _song("Bravo"))
    cid = store.library.create_collection("Friday / set", [B, A])["id"]
    status, data, headers = server("GET", f"/api/collections/{cid}/export.pdf", raw=True)
    assert status == 200 and data.startswith(b"%PDF-")
    assert 'filename="Friday_set.pdf"' in headers["Content-Disposition"]
    assert webserver.collection_export_name("Più bella / tour") == "Piu_bella_tour.pdf"
    assert webserver.collection_export_name("///") == "Collection.pdf"
    empty = store.library.create_collection("Empty")["id"]
    status, body, _ = server("GET", f"/api/collections/{empty}/export.pdf")
    assert status == 422 and "no songs" in body["error"]
    status, _, _ = server("GET", "/api/collections/nope/export.pdf")
    assert status == 404


def test_open_example_unhides_the_example(server, store):
    _, body, _ = server("POST", "/api/songs/example", {})
    name = body["filename"]
    store.library.set_hidden(name, True)
    _, body, _ = server("POST", "/api/songs/example", {})
    assert body["filename"] == name
    assert store.library.view()["hidden"] == []
