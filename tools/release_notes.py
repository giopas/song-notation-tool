#!/usr/bin/env python3
"""The title and text of a GitHub release, from CHANGELOG.md.

    python3 tools/release_notes.py v0.33.0                 # prints the title, then the notes
    python3 tools/release_notes.py v0.33.0 --title         # the title only
    python3 tools/release_notes.py v0.33.0 --out notes.md  # the notes into a file, the title printed

A release heading in CHANGELOG.md carries a few words on the main change
after the date:

    ## [0.33.0] — 2026-11-08 — Parts in tab or free text

The release is called "v0.33.0 (Parts in tab or free text)" and its text is
that section of the changelog, with a line pointing to the downloads and the
whole changelog. Without the few words, the release is called by its
version alone.

The release workflow (.github/workflows/release.yml) runs this when a v* tag
is pushed and publishes the release with this title and text, so there is
no draft to fill in by hand. Borrowed from QLC+ Swiss Knife.
"""
from __future__ import annotations

import argparse
import os
import re
import sys

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
REPO = "giopas/song-notation-tool"
_HEAD = re.compile(r"^## \[(?P<ver>[^\]]+)\]\s*(?:[—-]\s*(?P<date>\d{4}-\d{2}-\d{2}))?"
                   r"\s*(?:[—-]\s*(?P<title>.+?))?\s*$")


def section(version: str, text: str):
    """(few words, body) of one version's section; ValueError when it is not there."""
    v = version.lstrip("v")
    lines = text.splitlines()
    start, title = None, ""
    for i, line in enumerate(lines):
        m = _HEAD.match(line)
        if m and m.group("ver") == v:
            start, title = i + 1, (m.group("title") or "").strip()
            break
    if start is None:
        raise ValueError(f"CHANGELOG.md has no section for {v}")
    end = next((j for j in range(start, len(lines)) if lines[j].startswith("## ")), len(lines))
    body = "\n".join(lines[start:end]).strip()
    if not body:
        raise ValueError(f"the CHANGELOG.md section for {v} is empty")
    return title, body


def release(version: str, text: str):
    """(release title, release notes)."""
    tag = "v" + version.lstrip("v")
    words, body = section(tag, text)
    name = f"{tag} ({words})" if words else tag
    notes = (body + "\n\n---\n\n"
             "Downloads for macOS (Apple silicon and Intel), Windows and Linux, and "
             "their checksums, are below. The whole history is in the "
             f"[CHANGELOG](https://github.com/{REPO}/blob/main/CHANGELOG.md).\n")
    return name, notes


def main(argv=None) -> int:
    ap = argparse.ArgumentParser(description=__doc__.split("\n\n")[0])
    ap.add_argument("version", help="v0.33.0 or 0.33.0")
    ap.add_argument("--title", action="store_true", help="print the release title only")
    ap.add_argument("--out", help="write the notes to this file and print the title")
    ap.add_argument("--changelog", default=os.path.join(ROOT, "CHANGELOG.md"))
    a = ap.parse_args(argv)
    with open(a.changelog, encoding="utf-8") as f:
        text = f.read()
    try:
        name, notes = release(a.version, text)
    except ValueError as e:
        print(e, file=sys.stderr)
        return 1
    if a.out:
        with open(a.out, "w", encoding="utf-8") as f:
            f.write(notes)
        print(name)
    elif a.title:
        print(name)
    else:
        print(name)
        print()
        print(notes)
    return 0


if __name__ == "__main__":
    sys.exit(main())
