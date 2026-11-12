"""The pitch video and deck are made and attached by GitHub Actions, never by hand.

pitch.yml makes them with tools/make_pitch_video.py and attaches them to the
release; release.yml calls it after publishing. The README links to the files
of the latest release, and Update and restart must never pick them up.
"""
import importlib.util
import os
import re
import sys
from pathlib import Path

ROOT = Path(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
WF = ROOT / ".github" / "workflows"
sys.path.insert(0, str(ROOT))


def _tool():
    spec = importlib.util.spec_from_file_location("pitch", ROOT / "tools" / "make_pitch_video.py")
    m = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(m)
    return m


def test_release_calls_pitch_after_publishing():
    rel = (WF / "release.yml").read_text(encoding="utf-8")
    m = re.search(r"\n  pitch:\n((?:    .*\n|\s*\n)+)", rel + "\n")
    assert m, "release.yml has no pitch job"
    job = m.group(1)
    assert "needs: publish" in job
    assert "uses: ./.github/workflows/pitch.yml" in job
    assert "tag: ${{ github.ref_name }}" in job


def test_pitch_workflow_uses_the_strict_voice_and_attaches_both_files():
    text = (WF / "pitch.yml").read_text(encoding="utf-8")
    assert "--stills --tts edge --strict-voice" in text
    assert "gh release upload" in text
    assert "Song_Notation_Tool_Pitch.mp4 Song_Notation_Tool_Pitch.pdf" in text
    for trigger in ("  workflow_call:", "  workflow_dispatch:", "  push:",
                    '      - "tools/make_pitch_video.py"', '      - "screenshots/**"'):
        assert trigger in text, trigger


def test_update_and_restart_ignores_the_pitch_files():
    import update
    for name in ("Song_Notation_Tool_Pitch.mp4", "Song_Notation_Tool_Pitch.pdf"):
        assert not update._ASSET_RE.match(name)
    assert update.pick_asset([{"name": "Song_Notation_Tool_Pitch.mp4"}], "macos", "arm64") is None


def test_readme_links_the_latest_release_files_and_the_poster_exists():
    readme = (ROOT / "README.md").read_text(encoding="utf-8")
    m = _tool()
    assert m.FILE_STEM == "Song_Notation_Tool_Pitch"
    for ext in ("mp4", "pdf"):
        assert f"releases/latest/download/{m.FILE_STEM}.{ext}" in readme
    assert (ROOT / "screenshots" / "pitch-poster.png").is_file()


def test_every_zoom_point_names_a_picture_that_exists():
    m = _tool()
    for slide in m.SLIDES:
        for spec in slide.get("media", []):
            name = m.parse_media(spec)[0]
            assert m.find_media(name), f"{slide['title']}: {name}"
