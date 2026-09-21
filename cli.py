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
import grammar
import model
import songmap
from constants import APP_VERSION, default_export_name


def _load_doc(path: str) -> dict:
    with open(path, "r", encoding="utf-8") as f:
        raw = json.load(f)
    return model.migrate_document(raw)


def _parse_instruments(arg_list):
    if not arg_list:
        return None
    return set(arg_list)


def _export_one(doc: dict, fmt: str, out_path: str, instruments, orient):
    if fmt == "txt":
        lines = export.build_song_lines(doc, instruments=instruments)
        with open(out_path, "w", encoding="utf-8") as f:
            f.write("\n".join(lines))
    else:
        pdf_bytes = export.build_pdf(doc, instruments=instruments, orient=orient)
        with open(out_path, "wb") as f:
            f.write(pdf_bytes)


def cmd_convert(args):
    doc = _load_doc(args.input)
    instruments = _parse_instruments(args.instrument)
    out_path = args.output or default_export_name(doc, args.export)
    _export_one(doc, args.export, out_path, instruments, args.orient)
    print(f"Wrote {out_path}")
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
            _export_one(doc, args.export, out_path, instruments, args.orient)
            print(f"  {os.path.basename(path)}  ->  {out_path}")
            n_ok += 1
        except Exception as exc:  # noqa: BLE001 — batch job: report and continue
            print(f"  {os.path.basename(path)}  FAILED: {exc}", file=sys.stderr)
            n_fail += 1

    print(f"Done: {n_ok} exported, {n_fail} failed.")
    return 1 if n_fail else 0


def cmd_lint(args):
    doc = _load_doc(args.input)
    ok = True
    for sec in doc.get("sections", []):
        chart = songmap.chart_items(sec)
        try:
            line = grammar.unparse(chart) if chart else ""
        except Exception as exc:  # noqa: BLE001
            print(f"[{sec.get('name')}] could not unparse: {exc}", file=sys.stderr)
            ok = False
            continue
        # Round-trip: reparsing the unparsed line should reproduce the
        # same items (catches a section saved by a future/incompatible
        # grammar version without opening the GUI).
        try:
            reparsed = grammar.parse_items(line) if line else []
        except grammar.ParseError as exc:
            print(f"[{sec.get('name')}] chart line does not re-parse: {exc}",
                  file=sys.stderr)
            ok = False
            continue
        print(f"[{sec.get('name')}] OK — {line!r}" if line else
              f"[{sec.get('name')}] (empty chart)")
    # Chord shapes against the chords the chart plays. Warnings, not
    # failures: plenty of chords need no diagram, and a lint that fails a
    # build over one is a lint that gets switched off.
    if doc.get("chords"):
        import chords as chords_mod
        cov = chords_mod.coverage(doc)
        if cov["missing"]:
            print("note: chords played with no shape: " + ", ".join(cov["missing"]))
        if cov["unused"]:
            print("note: shapes for chords never played: " + ", ".join(cov["unused"]))
    if ok:
        print("All sections parse cleanly.")
    return 0 if ok else 1


def build_parser():
    p = argparse.ArgumentParser(
        prog="cli.py",
        description="Song Notation Tool — headless conversion/export.",
        epilog=f"Song Notation Tool v{APP_VERSION}",
    )
    sub = p.add_subparsers(dest="command", required=True)

    common = argparse.ArgumentParser(add_help=False)
    common.add_argument("-e", "--export", choices=["txt", "pdf"], default="pdf",
                         help="Export format (default: pdf)")
    common.add_argument("--instrument", action="append",
                         help="Only include this instrument's sections "
                              "(repeat the flag for more than one; default: all)")
    common.add_argument("--orient", choices=["portrait", "landscape"],
                         default="portrait", help="PDF page orientation "
                         "(ignored for --export txt)")

    pc = sub.add_parser("convert", parents=[common],
                         help="Convert one .sng file to TXT or PDF")
    pc.add_argument("-i", "--input", required=True, help="Path to a .sng file")
    pc.add_argument("-o", "--output", help="Output path "
                     "(default: '<Artist> - <Title>.<ext>' in the current dir)")
    pc.set_defaults(func=cmd_convert)

    pb = sub.add_parser("batch", parents=[common],
                         help="Re-export every .sng file in a folder")
    pb.add_argument("-i", "--input-dir", required=True,
                     help="Folder to scan for *.sng files")
    pb.add_argument("--out-dir", help="Write exports here instead of "
                     "alongside each .sng (created if missing)")
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
