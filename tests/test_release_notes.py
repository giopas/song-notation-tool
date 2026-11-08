"""A release is published straight from its CHANGELOG.md section (v0.33.0).

The release workflow names the release and writes its text from the
changelog, then publishes it: no draft to fill in by hand. These tests
check the script that does it, that the version being released has a
section with its few words, and that the workflow publishes only after
every build has succeeded.
"""
import os
import sys

import pytest

ROOT = os.path.join(os.path.dirname(__file__), "..")
sys.path.insert(0, os.path.join(ROOT, "tools"))
sys.path.insert(0, ROOT)

import release_notes  # noqa: E402
from constants import APP_VERSION  # noqa: E402

SAMPLE = """# Changelog

## [Unreleased]

## [1.2.0] — 2026-12-01 — Bigger notes

Why this version.

### Added

- Something.

## [1.1.0] — 2026-11-01

Older.
"""


def test_title_from_the_few_words_after_the_date():
    name, notes = release_notes.release("v1.2.0", SAMPLE)
    assert name == "v1.2.0 (Bigger notes)"
    assert notes.startswith("Why this version.")
    assert "- Something." in notes and "Older." not in notes
    assert "CHANGELOG" in notes


def test_a_heading_without_words_is_called_by_its_version():
    name, notes = release_notes.release("1.1.0", SAMPLE)
    assert name == "v1.1.0" and notes.startswith("Older.")


def test_a_missing_version_is_an_error():
    with pytest.raises(ValueError):
        release_notes.release("v9.9.9", SAMPLE)


def test_the_current_version_has_a_section_with_its_few_words():
    text = open(os.path.join(ROOT, "CHANGELOG.md"), encoding="utf-8").read()
    words, body = release_notes.section(APP_VERSION, text)
    assert words, f"add a few words after the date in the {APP_VERSION} heading"
    assert body


def test_the_workflow_publishes_from_the_changelog():
    wf = os.path.join(ROOT, ".github", "workflows", "release.yml")
    if not os.path.exists(wf):
        pytest.skip("release.yml is copied into .github/workflows by hand")
    text = open(wf, encoding="utf-8").read()
    assert "tools/release_notes.py" in text
    # Files go up while the release is hidden, then it is published.
    assert "--draft=false" in text and "--latest" in text
    assert text.index("gh release upload") < text.index("--draft=false")
    assert "needs.build.result == 'success'" in text
