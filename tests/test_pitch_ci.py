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


def test_the_readme_poster_opens_the_page_that_plays_the_video():
    readme = (ROOT / "README.md").read_text(encoding="utf-8")
    assert "(screenshots/pitch-poster.png)](https://giopas.github.io/song-notation-tool/)" in readme


def test_the_page_workflow_runs_from_main_and_pitch_starts_it():
    page = (WF / "pitch-page.yml").read_text(encoding="utf-8")
    assert "  workflow_dispatch:" in page and "push:" not in page
    for need in ("pages: write", "id-token: write", "name: github-pages",
                 "actions/deploy-pages@", "cp tools/pitch_page.html site/index.html",
                 "-p 'Song_Notation_Tool_Pitch.*'"):
        assert need in page, need
    pitch = (WF / "pitch.yml").read_text(encoding="utf-8")
    assert "gh workflow run pitch-page.yml" in pitch and "--ref main" in pitch
    assert "actions: write" in pitch
    rel = (WF / "release.yml").read_text(encoding="utf-8")
    job = rel[rel.index("\n  pitch:\n"):]
    assert "actions: write" in job and "contents: write" in job


def test_the_page_plays_the_files_the_workflow_puts_beside_it():
    html = (ROOT / "tools" / "pitch_page.html").read_text(encoding="utf-8")
    m = _tool()
    for name in (f"{m.FILE_STEM}.mp4", f"{m.FILE_STEM}.pdf", "icon.png", "slide-01.png"):
        assert f'"{name}"' in html, name
    assert "<video" in html and "playsinline" in html


def test_a_release_called_run_is_not_mistaken_for_a_push_to_main():
    """Called by release.yml, GITHUB_EVENT_NAME is the tag's "push": the skip
    for a release commit pushed to main must look at the branch and the input."""
    text = (WF / "pitch.yml").read_text(encoding="utf-8")
    line = next(l for l in text.splitlines() if "git tag --points-at HEAD" in l)
    assert '-z "$IN_TAG"' in line
    assert '"$GITHUB_REF" = "refs/heads/main"' in line
    assert "GITHUB_EVENT_NAME" not in line
    paths = text[text.index("    paths:"):].splitlines()
    assert paths[1].strip() == '- ".github/workflows/pitch.yml"'
