#!/usr/bin/env python3
"""
cli.py — Song Notation Tool headless CLI.

No Tkinter import, no window — built entirely on the pure-data layers
(model, export) so a whole folder of songs can be regenerated from a
script or CI job. See "Enhancement 2 — browser and headless use" in the
UX & Multi-Platform Enhancement Spec, phased-roadmap step 3.

Examples
--------
Convert one song to PDF:

    ./cli.py convert -i song.sng -e pdf -o song.pdf

Convert one song to TXT, only the bass parts, landscape orientation
(orientation only matters for PDF):

    ./cli.py convert -i song.sng -e txt --instrument "Bass (4-string)"

Regenerate every .sng in a folder as PDF, e.g. before a gig or after a
formatting change (writes alongside each .sng unless --out-dir is given):

    ./cli.py batch -i songs/ -e pdf

    ./cli.py batch -i songs/ -e txt --out-dir exports/

Show what a song's chart lines currently parse to (a quick way to
sanity-check a .sng file, or a file emitted by another tool, without
opening the GUI):

    ./cli.py lint -i song.sng
"""

from __future__ import annotations

import argparse
import glob
import json
import os
import sys

import export
import model
from constants import APP_VERSION, default_export_name


def _load_doc(path: str) -> dict:
    with open(path, "r", encoding="utf-8") as f:
        raw = json.load(f)
    return model.migrate_document(raw)


def _parse_instruments(arg_list):
    if not arg_list:
        return None
    return set(arg_list)


def _export_one(doc: dict, fmt: str, out_path: str, instruments, orient,
                force: bool = False) -> str:
    """Check, build and write one export. Returns the path written.

    An existing file is kept and the export goes to the next free name
    (`song_v2.pdf`) unless `force` is set. Warnings from the check are
    printed; errors raise `export.ExportBlocked` and nothing is written.
    """
    result = export.export_to_file(doc, fmt, out_path, instruments=instruments,
                                   orient=orient, overwrite=force)
    for w in result["warnings"]:
        print(f"note: {w}")
    return result["path"]


def cmd_convert(args):
    doc = _load_doc(args.input)
    instruments = _parse_instruments(args.instrument)
    out_path = args.output or default_export_name(doc, args.export)
    try:
        written = _export_one(doc, args.export, out_path, instruments,
                              args.orient, force=args.force)
    except export.ExportBlocked as exc:
        for err in exc.errors:
            print(err, file=sys.stderr)
        print("Not exported: fix the chart lines above first.", file=sys.stderr)
        return 1
    if os.path.abspath(written) != os.path.abspath(out_path):
        print(f"{out_path} already exists, so it was kept.")
    print(f"Wrote {written}")
    return 0


def cmd_batch(args):
    pattern = os.path.join(args.input_dir, "*.sng")
    paths = sorted(glob.glob(pattern))
    if not paths:
        print(f"No .sng files found in {args.input_dir}", file=sys.stderr)
        return 1

    instruments = _parse_instruments(args.instrument)
    out_dir = args.out_dir or args.input_dir
    os.makedirs(out_dir, exist_ok=True)

    n_ok, n_fail = 0, 0
    for path in paths:
        try:
            doc = _load_doc(path)
            stem = os.path.splitext(os.path.basename(path))[0]
            out_path = os.path.join(out_dir, f"{stem}.{args.export}")
            written = _export_one(doc, args.export, out_path, instruments,
                                  args.orient, force=args.force)
            print(f"  {os.path.basename(path)}  ->  {written}")
            n_ok += 1
        except export.ExportBlocked as exc:
            print(f"  {os.path.basename(path)}  NOT EXPORTED: "
                  + "; ".join(exc.errors), file=sys.stderr)
            n_fail += 1
        except Exception as exc:  # noqa: BLE001 — batch job: report and continue
            print(f"  {os.path.basename(path)}  FAILED: {exc}", file=sys.stderr)
            n_fail += 1

    print(f"Done: {n_ok} exported, {n_fail} failed.")
    return 1 if n_fail else 0


def cmd_lint(args):
    """The same check every export runs first (`export.check_document`)."""
    doc = _load_doc(args.input)
    report = export.check_document(doc)
    for sec in report["sections"]:
        if sec["error"]:
            print(sec["error"], file=sys.stderr)
        elif sec["line"]:
            print(f"[{sec['name']}] OK: {sec['line']!r}")
        else:
            print(f"[{sec['name']}] (empty chart)")
    # Chord shapes against the chords the chart plays. Warnings, not
    # failures: plenty of chords need no diagram, and a lint that fails a
    # build over one is a lint that gets switched off.
    for w in report["warnings"]:
        print(f"note: {w}")
    if report["ok"]:
        print("All sections parse cleanly.")
    return 0 if report["ok"] else 1


def build_parser():
    p = argparse.ArgumentParser(
        prog="cli.py",
        description="Song Notation Tool: convert and export songs from the command line.",
        epilog=f"Song Notation Tool v{APP_VERSION}",
    )
    sub = p.add_subparsers(dest="command", required=True)

    common = argparse.ArgumentParser(add_help=False)
    common.add_argument("-e", "--export", choices=["txt", "pdf"], default="pdf",
                         help="Export format (default: pdf)")
    common.add_argument("--instrument", action="append",
                         help="Only include this instrument's sections "
                              "(repeat the flag for more than one; default: all)")
    common.add_argument("--force", action="store_true",
                         help="Replace an existing output file. Without it an "
                              "existing file is kept and the export gets the "
                              "next free name (song_v2.pdf)")
    common.add_argument("--orient", choices=["portrait", "landscape"],
                         default="portrait", help="PDF page orientation "
                         "(ignored for --export txt)")

    pc = sub.add_parser("convert", parents=[common],
                         help="Convert one .sng file to TXT or PDF")
    pc.add_argument("-i", "--input", required=True, help="Path to a .sng file")
    pc.add_argument("-o", "--output", help="Output path "
                     "(default: '<Artist> - <Title>.<ext>' in the current folder)")
    pc.set_defaults(func=cmd_convert)

    pb = sub.add_parser("batch", parents=[common],
                         help="Re-export every .sng file in a folder")
    pb.add_argument("-i", "--input-dir", required=True,
                     help="Folder to scan for *.sng files")
    pb.add_argument("--out-dir", help="Write exports here instead of "
                     "next to each .sng (created if missing)")
    pb.set_defaults(func=cmd_batch)

    pl = sub.add_parser("lint", help="Parse every section's chart line and "
                         "report errors, without opening the GUI")
    pl.add_argument("-i", "--input", required=True, help="Path to a .sng file")
    pl.set_defaults(func=cmd_lint)

    return p


def main(argv=None):
    parser = build_parser()
    args = parser.parse_args(argv)
    return args.func(args)


if __name__ == "__main__":
    sys.exit(main())
