"""One writer for every export (export.py), exports never overwrite, and
every export is checked first. The command line is tested through a real
subprocess and the web server through its real HTTP routes."""
import datetime
import http.server
import json
import os
import re
import subprocess
import sys
import threading
import urllib.error
import urllib.request

ROOT = os.path.abspath(os.path.join(os.path.dirname(__file__), ".."))
sys.path.insert(0, ROOT)

import pytest

import export
import grammar
import model

DAY = datetime.date(2026, 1, 2)


def _doc(chords=None):
    doc = model.new_document(title="Glass Harbour", artist="The Placeholders",
                             key="B", time="4/4", bpm="")
    sec = model.new_section("intro", "Intro", "Intro", instrument="Bass (4-string)")
    sec["items"] = [model.make_token("B", 7), model.make_token("E", None)]
    doc["sections"].append(sec)
    if chords is not None:
        doc["chords"] = chords
    return doc


def _break_the_grammar(monkeypatch):
    def refuse(line):
        raise grammar.ParseError("simulated: this line no longer parses")
    monkeypatch.setattr(grammar, "parse_items", refuse)


# ── item 3: one writer ────────────────────────────────────────────────────

def test_txt_bytes_are_utf8_with_unix_endings_and_a_final_newline():
    data = export.export_bytes(_doc(), "txt")
    text = data.decode("utf-8")
    assert "\r" not in text and text.endswith("\n") and not text.endswith("\n\n")
    assert text == "\n".join(export.build_song_lines(_doc())) + "\n"


def test_pdf_bytes_are_build_pdf():
    assert export.export_bytes(_doc(), "pdf", date=DAY) == export.build_pdf(_doc(), date=DAY)


def test_unknown_format_is_refused():
    with pytest.raises(ValueError):
        export.export_bytes(_doc(), "docx")


def test_front_ends_never_write_export_files_themselves():
    """Only export.py writes TXT and PDF files. The one write left in each
    front end is the .sng song file itself."""
    def opens_for_writing(name):
        src = open(os.path.join(ROOT, name), encoding="utf-8").read()
        return re.findall(r"open\([^)]*[\"']w[b]?[\"']", src)
    assert opens_for_writing("cli.py") == []
    assert len(opens_for_writing("song_writer.py")) == 1   # the .sng save
    assert len(opens_for_writing("webserver.py")) == 2     # SongStore's temp file,
                                                           # and --smoke's update feed
    for name in ("cli.py", "song_writer.py", "webserver.py"):
        src = open(os.path.join(ROOT, name), encoding="utf-8").read()
        assert "export_to_file(" in src


# ── item 4: never overwrite ───────────────────────────────────────────────

def test_an_existing_file_is_kept_and_the_export_gets_the_next_name(tmp_path):
    target = tmp_path / "song.pdf"
    target.write_bytes(b"mine")
    first = export.write_export(str(target), b"one")
    second = export.write_export(str(target), b"two")
    assert target.read_bytes() == b"mine"
    assert os.path.basename(first) == "song_v2.pdf"
    assert os.path.basename(second) == "song_v3.pdf"
    assert (tmp_path / "song_v3.pdf").read_bytes() == b"two"


def test_a_free_name_is_used_as_is():
    assert export.next_free_path("/nonexistent/folder/x.pdf") == "/nonexistent/folder/x.pdf"


def test_vN_names_continue_from_their_number(tmp_path):
    (tmp_path / "song_v2.pdf").write_bytes(b"x")
    assert os.path.basename(export.next_free_path(str(tmp_path / "song_v2.pdf"))) == "song_v3.pdf"


def test_overwrite_replaces_only_when_asked(tmp_path):
    target = tmp_path / "song.txt"
    target.write_bytes(b"old")
    assert export.write_export(str(target), b"new", overwrite=True) == str(target)
    assert target.read_bytes() == b"new"


def test_no_temporary_file_is_left_behind(tmp_path):
    export.write_export(str(tmp_path / "a.txt"), b"x")
    export.write_export(str(tmp_path / "a.txt"), b"y")
    export.write_export(str(tmp_path / "a.txt"), b"z", overwrite=True)
    assert sorted(os.listdir(tmp_path)) == ["a.txt", "a_v2.txt"]


# ── item 5: check before export ───────────────────────────────────────────

def test_a_good_song_passes_the_check():
    report = export.check_document(_doc())
    assert report["ok"] and report["errors"] == []
    assert report["sections"][0]["line"] == "7B E"


def test_a_failing_check_blocks_the_export_and_writes_nothing(tmp_path, monkeypatch):
    _break_the_grammar(monkeypatch)
    with pytest.raises(export.ExportBlocked) as info:
        export.export_to_file(_doc(), "pdf", str(tmp_path / "x.pdf"))
    assert "Intro" in info.value.errors[0]
    assert os.listdir(tmp_path) == []


def test_chord_shape_warnings_never_block(tmp_path):
    doc = _doc(chords=[{"name": "C", "shape": "x32010"}])
    result = export.export_to_file(doc, "txt", str(tmp_path / "x.txt"))
    assert os.path.exists(result["path"])
    assert any("never played" in w for w in result["warnings"])


# ── through the command line ──────────────────────────────────────────────

def _cli(*args, cwd):
    return subprocess.run([sys.executable, os.path.join(ROOT, "cli.py"), *args],
                          cwd=cwd, capture_output=True, text=True, timeout=60)


def _sng(tmp_path, doc=None):
    path = tmp_path / "song.sng"
    path.write_text(json.dumps(doc or _doc()), encoding="utf-8")
    return path


def test_cli_convert_keeps_an_existing_file_unless_forced(tmp_path):
    sng = _sng(tmp_path)
    out = tmp_path / "out.txt"
    out.write_text("keep me", encoding="utf-8")

    r = _cli("convert", "-i", str(sng), "-e", "txt", "-o", str(out), cwd=tmp_path)
    assert r.returncode == 0, r.stderr
    assert out.read_text(encoding="utf-8") == "keep me"
    assert (tmp_path / "out_v2.txt").read_bytes() == export.export_bytes(_doc(), "txt")
    assert "kept" in r.stdout

    r = _cli("convert", "-i", str(sng), "-e", "txt", "-o", str(out), "--force",
             cwd=tmp_path)
    assert r.returncode == 0, r.stderr
    assert out.read_bytes() == export.export_bytes(_doc(), "txt")
    assert not (tmp_path / "out_v3.txt").exists()


def test_cli_batch_writes_next_to_each_song_without_overwriting(tmp_path):
    _sng(tmp_path)
    assert _cli("batch", "-i", str(tmp_path), "-e", "txt", cwd=tmp_path).returncode == 0
    assert _cli("batch", "-i", str(tmp_path), "-e", "txt", cwd=tmp_path).returncode == 0
    assert (tmp_path / "song.txt").exists() and (tmp_path / "song_v2.txt").exists()


def test_cli_lint_and_convert_share_the_check(tmp_path):
    sng = _sng(tmp_path, _doc(chords=[{"name": "C", "shape": "x32010"}]))
    r = _cli("lint", "-i", str(sng), cwd=tmp_path)
    assert r.returncode == 0
    assert "[Intro] OK: '7B E'" in r.stdout
    assert "note: shapes for chords never played: C" in r.stdout
    r = _cli("convert", "-i", str(sng), "-e", "pdf", cwd=tmp_path)
    assert r.returncode == 0 and "note: shapes for chords never played" in r.stdout


# ── through the web server's HTTP routes ──────────────────────────────────

@pytest.fixture
def server(tmp_path):
    import webserver
    webserver.Handler.store = webserver.SongStore(str(tmp_path))
    httpd = http.server.ThreadingHTTPServer(("127.0.0.1", 0), webserver.Handler)
    thread = threading.Thread(target=httpd.serve_forever, daemon=True)
    thread.start()
    yield f"http://127.0.0.1:{httpd.server_address[1]}"
    httpd.shutdown()
    httpd.server_close()


def _post(url, body):
    req = urllib.request.Request(url, data=json.dumps(body).encode("utf-8"),
                                 headers={"Content-Type": "application/json"})
    return urllib.request.urlopen(req, timeout=30)


def test_web_download_is_the_same_bytes_as_the_cli(server):
    with _post(server + "/api/export.txt", {"doc": _doc()}) as res:
        assert res.read() == export.export_bytes(_doc(), "txt")
        assert res.headers.get("X-Export-Warnings") is None


def test_web_download_reports_warnings(server):
    doc = _doc(chords=[{"name": "C", "shape": "x32010"}])
    with _post(server + "/api/export.pdf", {"doc": doc}) as res:
        assert res.read().startswith(b"%PDF")
        assert any("never played" in w for w in json.loads(res.headers["X-Export-Warnings"]))


def test_web_download_is_refused_when_the_check_fails(server, monkeypatch):
    _break_the_grammar(monkeypatch)
    with pytest.raises(urllib.error.HTTPError) as info:
        _post(server + "/api/export.pdf", {"doc": _doc()})
    assert info.value.code == 422
    body = json.loads(info.value.read())
    assert body["ok"] is False and "Intro" in body["errors"][0]
