"""tools/shots.py takes exactly the pictures the README shows.

A screenshot added to the README without a step in tools/shots.py would
never be retaken, and one the tool takes but the README does not use is
a file nobody looks at.
"""
import os
import re
import sys

ROOT = os.path.join(os.path.dirname(__file__), "..")
sys.path.insert(0, os.path.join(ROOT, "tools"))

import shots  # noqa: E402  (Playwright is imported only when pictures are taken)


def _readme_pictures():
    readme = open(os.path.join(ROOT, "README.md"), encoding="utf-8").read()
    # The video's poster is drawn by make_pitch_video.py --poster, not
    # taken from the app, so tools/shots.py has no step for it.
    return set(re.findall(r"screenshots/([\w-]+\.png)", readme)) - {"pitch-poster.png"}


def test_the_tool_and_the_readme_name_the_same_pictures():
    assert {name for name, _what in shots.PICTURES} == _readme_pictures()


def test_every_picture_is_in_the_repo():
    for name, _what in shots.PICTURES:
        assert os.path.exists(os.path.join(ROOT, "screenshots", name)), name
