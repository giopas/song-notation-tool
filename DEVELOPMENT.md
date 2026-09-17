# Development workflow

This project is built through iterative, AI-assisted ("vibecoded")
development, working from the same handwritten reference sheets each time.
This file exists so a new session — human or AI — starts from a consistent
brief instead of re-deriving context from the code.

## The actual goal

The target output is a one-page stage chart, matching the vocabulary the
author already uses on paper: fret numbers over chord symbols, repeat
counts at every level, named riffs referenced by name, back-references
("same as the verse above"), and free-form annotations. See
`DESIGN_v0_16.md` section 1 for the full vocabulary and the reference case
this repo audits against.

Any change should be checked against that goal, not just against whether
the code runs. A feature that makes the tab grid more powerful but doesn't
shrink the output toward a one-page chart is not the priority.

## Starting a session

1. Read the current `DESIGN_v0_...md` for the version in progress, if one
   exists. It is the source of truth for what's being built and why —
   more so than the code, if they disagree mid-implementation.
2. Read `CHANGELOG.md` and `ROADMAP.md` to see what's already shipped and
   what's explicitly deferred. Don't re-propose deferred items without
   flagging that they were deliberately deferred.
3. Run the test suite (`python3 -m pytest tests/`) before changing
   anything, so failures introduced during the session are attributable.

## Build order

Pure-data modules first, UI last — this project's own build order
(`model.py` → `grammar.py` → `transpose.py` → `render.py` → wire into the
exporters → minimal UI → repo cleanup) has held up well and is worth
reusing: logic with no Tkinter dependency is fully unit-testable without a
display, which this environment often doesn't have.

## Conventions

- No dependencies beyond the Python standard library, plus `pytest` for
  the test suite (dev-only), and an entirely optional `pywebview` for
  `webserver.py`'s native-window launch mode (see
  `requirements-optional.txt`) — never required to run anything.
- The version lives in `APP_VERSION` in `constants.py` — nowhere else.
  Do not version the filename.
- A bug fix belongs in the changelog under "Fixed" even if it was
  discovered incidentally while building something else.
- Prefer fixing a shared root cause (e.g. one canonical `ENHARMONIC` table)
  over patching each symptom separately.

## Releasing

1. Update `CHANGELOG.md` — a short "why" paragraph at the top of the
   version's entry, then Added/Changed/Removed/Deferred as usual. This
   is the single source of release history now; there's no separate
   `RELEASE_NOTES_vX.Y.Z.md` file per release any more (there used to
   be, for v0.16.0 and v0.17.0 — folded back into `CHANGELOG.md` and
   removed once the app's pace made maintaining two documents of the
   same story per release more overhead than it was worth).
2. Tag `vX.Y.Z` and cut a GitHub Release from the tag, using that
   version's `CHANGELOG.md` section as the release body.
