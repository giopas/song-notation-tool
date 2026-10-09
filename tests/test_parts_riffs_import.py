"""v0.30: parts for other instruments, the riff library in the browser,
and reopening a song from its PDF."""
import base64
import json
import os
import sys
import threading
import urllib.error
import urllib.request

import pytest

sys.path.insert(0, os.path.join(os.path.dirname(__file__), ".."))

import chords  # noqa: E402
import export  # noqa: E402
import grammar  # noqa: E402
import model  # noqa: E402
import songmap  # noqa: E402

BASS, GUITAR = "Bass (4-string)", "Guitar (6-string)"


def song_with_a_part():
    doc = model.new_document("Two Parts", "Band")
    sec = model.new_section("s1", "Verse", instrument=BASS)
    sec["items"], _ = grammar.parse("5A 7D")
    gtr, _ = grammar.parse("Am C G")
    sec["parts"] = [model.make_part("p1", GUITAR, gtr)]
    doc["sections"].append(sec)
    return doc


# ── parts ────────────────────────────────────────────────────────────────

def test_a_part_prints_under_its_section():
    lines = export.build_song_lines(song_with_a_part())
    text = "\n".join(lines)
    assert GUITAR + ":" in text
    assert text.index("5") < text.index("Am")


def test_layout_can_print_only_one_instrument():
    doc = song_with_a_part()
    doc["print_instruments"] = [GUITAR]
    text = "\n".join(export.build_song_lines(doc))
    assert "Am" in text and "  5  7" not in text
    doc["print_instruments"] = [BASS]
    text = "\n".join(export.build_song_lines(doc))
    assert "Am" not in text


def test_a_choice_that_matches_nothing_prints_everything():
    doc = song_with_a_part()
    doc["print_instruments"] = ["Ukulele"]
    assert export.instruments_to_print(doc) == {BASS, GUITAR}


def test_a_song_with_parts_makes_a_pdf():
    assert export.build_pdf(song_with_a_part()).startswith(b"%PDF")


def test_chords_in_parts_count_as_played():
    assert {"Am", "C", "G"} <= set(chords.used_symbols(song_with_a_part()))


def test_an_unparsable_part_blocks_the_export():
    doc = song_with_a_part()
    doc["sections"][0]["parts"][0]["items"] = [{"kind": "nonsense"}]
    assert not export.check_document(doc)["ok"]


# ── riffs ────────────────────────────────────────────────────────────────

def test_rename_a_riff_updates_every_reference():
    doc = song_with_a_part()
    songmap.create_riff(doc, "Main", grammar.parse("5A 7A")[0])
    doc["sections"][0]["items"].append(model.make_block_ref("Main"))
    doc["sections"][0]["parts"][0]["items"].append(model.make_block_ref("Main", 2))
    songmap.rename_riff(doc, "Main", "Hook")
    assert "Hook" in doc["blocks"] and "Main" not in doc["blocks"]
    sec = doc["sections"][0]
    assert sec["items"][-1]["block"] == "Hook"
    assert sec["parts"][0]["items"][-1]["block"] == "Hook"


@pytest.mark.parametrize("name, words", [
    ("two words", "one word"), ("Am", "chord"), ("mf", "dynamic"),
    ("Main", "already"),
])
def test_riff_names_are_checked(name, words):
    doc = model.new_document()
    songmap.create_riff(doc, "Main")
    with pytest.raises(songmap.RiffError, match=words):
        songmap.check_riff_name(doc, name)


def test_a_riff_in_use_cannot_be_deleted():
    doc = song_with_a_part()
    songmap.create_riff(doc, "Main")
    doc["sections"][0]["items"].append(model.make_block_ref("Main"))
    with pytest.raises(songmap.RiffError, match="still used"):
        songmap.delete_riff(doc, "Main")
    doc["sections"][0]["items"].pop()
    songmap.delete_riff(doc, "Main")
    assert "Main" not in doc["blocks"]


# ── the song inside its PDF ──────────────────────────────────────────────

def test_a_pdf_reopens_as_the_same_song():
    doc = song_with_a_part()
    pdf = export.build_pdf(doc)
    assert b"/EmbeddedFiles" in pdf
    assert export.song_from_bytes(pdf, "x.pdf") == model.migrate_document(doc)


def test_the_editor_only_fields_are_not_embedded():
    doc = song_with_a_part()
    doc["sections"][0]["chart_line"] = "5A 7D"
    assert "chart_line" not in json.loads(export.song_bytes(doc))["sections"][0]


def test_a_pdf_from_elsewhere_says_so():
    with pytest.raises(export.SongImportError, match="no song inside"):
        export.song_from_bytes(b"%PDF-1.4\n%%EOF\n", "other.pdf")


def test_text_that_is_not_a_song_says_so():
    with pytest.raises(export.SongImportError, match="not a song file"):
        export.song_from_bytes(b"hello", "notes.txt")


def test_a_sng_file_loads_too(tmp_path):
    path = tmp_path / "a.sng"
    path.write_text(json.dumps(song_with_a_part()), encoding="utf-8")
    assert export.load_song_file(str(path))["meta"]["title"] == "Two Parts"


# ── the HTTP routes ──────────────────────────────────────────────────────

@pytest.fixture()
def server(tmp_path):
    import webserver
    old = getattr(webserver.Handler, "store", None)
    webserver.Handler.store = webserver.SongStore(str(tmp_path / "songs"))
    httpd = webserver._ThreadingHTTPServer(("127.0.0.1", 0), webserver.Handler)
    threading.Thread(target=httpd.serve_forever, daemon=True).start()
    base = f"http://127.0.0.1:{httpd.server_address[1]}"

    def post(path, body):
        req = urllib.request.Request(base + path, data=json.dumps(body).encode(),
                                     headers={"Content-Type": "application/json"})
        try:
            with urllib.request.urlopen(req, timeout=10) as r:
                return r.status, json.loads(r.read())
        except urllib.error.HTTPError as e:
            return e.code, json.loads(e.read())

    yield post, tmp_path / "songs"
    httpd.shutdown()
    httpd.server_close()
    webserver.Handler.store = old


def test_import_a_pdf_makes_a_new_song_file(server):
    post, folder = server
    pdf = export.build_pdf(song_with_a_part())
    body = {"filename": "Band - Two Parts.pdf",
            "data": base64.b64encode(pdf).decode()}
    status, res = post("/api/import", body)
    assert status == 200 and res["filename"] == "Band - Two Parts.sng"
    assert res["doc"]["sections"][0]["parts"][0]["chart_line"] == "Am C G"
    # A second import never overwrites the first.
    status, res = post("/api/import", body)
    assert res["filename"] == "Band - Two Parts (2).sng"
    assert sorted(os.listdir(folder)) == ["Band - Two Parts (2).sng",
                                          "Band - Two Parts.sng"]


def test_import_refuses_a_file_with_no_song(server):
    post, folder = server
    status, res = post("/api/import", {"filename": "x.pdf",
                                       "data": base64.b64encode(b"%PDF-1.4").decode()})
    assert status == 400 and "no song inside" in res["error"]


def test_riff_route_create_rename_and_list(server):
    post, _ = server
    doc = song_with_a_part()
    status, res = post("/api/riffs", {"op": "create", "doc": doc,
                                      "name": "Main", "line": "5A 7A"})
    assert status == 200 and res["riffs"][0]["line"] == "5A 7A"
    doc = res["doc"]
    doc["sections"][0]["items"].append(model.make_block_ref("Main"))
    status, res = post("/api/riffs", {"op": "rename", "doc": doc,
                                      "id": "Main", "name": "Hook"})
    assert status == 200
    assert res["riffs"] == [{"id": "Hook", "line": "5A 7A", "used_in": ["Verse"]}]
    assert res["doc"]["sections"][0]["chart_line"].endswith("Hook")


def test_riff_route_reports_mistakes(server):
    post, _ = server
    status, res = post("/api/riffs", {"op": "create", "doc": model.new_document(),
                                      "name": "Am"})
    assert status == 400 and "chord" in res["error"]
