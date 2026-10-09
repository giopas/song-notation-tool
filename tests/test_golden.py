"""Golden files: invented songs, one per layout feature, rebuilt and
compared with the approved TXT and PDF in tests/golden/.

TXT is compared byte for byte. PDF is compared on everything except how
zlib packed the page streams: the streams are compared after inflating
them, because different zlib builds (zlib-ng in some Python 3.14 builds,
for one) pack the same content into different bytes. The drawing itself
must match exactly.

To change an expected file on purpose: python3 tools/make_golden.py, then
look at the new PDFs before committing them.
"""
import json
import os
import re
import sys
import zlib

ROOT = os.path.abspath(os.path.join(os.path.dirname(__file__), ".."))
sys.path.insert(0, ROOT)
sys.path.insert(0, os.path.join(ROOT, "tools"))

import pytest

import make_golden
import model

STREAM = re.compile(rb"<< /Length \d+ /Filter /FlateDecode >>\nstream\n(.*?)\nendstream",
                    re.S)


def normalise_pdf(data: bytes) -> bytes:
    """The PDF with each page stream inflated, and lengths and the
    cross-reference table (byte offsets) left out."""
    body = STREAM.sub(lambda m: b"<< stream >>\n" + zlib.decompress(m.group(1)), data)
    return body.split(b"\nxref\n")[0]


def _read(name):
    with open(os.path.join(make_golden.GOLDEN_DIR, name), "rb") as f:
        return f.read()


@pytest.mark.parametrize("name", sorted(make_golden.SONGS))
def test_golden_song(name):
    build, orient = make_golden.SONGS[name]
    stored = json.loads(_read(f"{name}.sng"))
    # The stored .sng is what the expected files were made from; it must
    # still be what the builder makes, and still open in this version.
    assert stored == json.loads(json.dumps(build()))
    out = make_golden.outputs(model.migrate_document(stored), orient)
    assert out["txt"] == _read(f"{name}.txt"), f"{name}.txt changed"
    assert normalise_pdf(out["pdf"]) == normalise_pdf(_read(f"{name}.pdf")), \
        f"{name}.pdf changed"


def test_git_keeps_the_golden_files_as_they_are():
    """A Windows checkout turns text files into CRLF unless told not to,
    and the golden TXT files are compared byte for byte."""
    attrs = open(os.path.join(ROOT, ".gitattributes"), encoding="utf-8").read()
    assert "tests/golden/* -text" in attrs
    for name in make_golden.SONGS:
        assert b"\r\n" not in _read(f"{name}.txt")


def test_every_golden_song_is_invented():
    """No real band, song or lyric goes into the repository."""
    for name in make_golden.SONGS:
        doc = json.loads(_read(f"{name}.sng"))
        assert doc["meta"]["artist"] == make_golden.BAND


def _repack(pdf, change=lambda content: content, level=1):
    return STREAM.sub(
        lambda m: b"<< /Length 0 /Filter /FlateDecode >>\nstream\n"
        + zlib.compress(change(zlib.decompress(m.group(1))), level)
        + b"\nendstream", pdf)


def test_normalising_ignores_how_zlib_packed_the_streams():
    pdf = _read("basics.pdf")
    assert _repack(pdf) != pdf
    assert normalise_pdf(_repack(pdf)) == normalise_pdf(pdf)


def test_normalising_still_sees_a_change_in_the_drawing():
    pdf = _read("basics.pdf")
    changed = _repack(pdf, lambda c: c.replace(b"PAPER LANTERNS", b"PAPER LANTERNZ"))
    assert normalise_pdf(changed) != normalise_pdf(pdf)
