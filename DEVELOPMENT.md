# Development workflow

This project is built through AI-assisted ("vibecoded") development. This
file is the working agreement: how a piece of work starts, how it is
built, tested, documented and released. It exists so a new session, human
or AI, starts from the same brief. The plan and the decisions in force are
in `WORKPLAN.md`.

## Roles

- **giopas decides.** Scope, order of work, release timing and anything
  that changes how the app behaves. Decisions go into the table in
  `WORKPLAN.md` with the date, so they are not asked twice.
- **Claude builds and checks.** Code, tests, documentation and release
  material, and it says plainly what it could not verify.
- **giopas pushes.** Claude commits locally and gives the exact commands.
- **giopas tests on real systems.** Claude runs the automated tests. Live
  checks (the desktop app, the native window, a printed page, Windows,
  Linux) are his, and each one stays on the "what is left" list until he
  reports back.
- **Live feedback goes into the docs.** When a step in the README or the
  wiki loses him during a test, that is a bug in the guide, fixed in the
  same session.

## The actual goal

The target output is a one-page stage chart, matching the vocabulary the
author already uses on paper: fret numbers over chord symbols, repeat
counts at every level, named riffs referenced by name, back-references
("same as the verse above") and free annotations. See `DESIGN_v0_16.md`
section 1 for the vocabulary and the reference case.

A change is checked against that goal, not only against whether the code
runs. A feature that makes the tab grid more powerful but does not shrink
the output toward a one-page chart is not the priority.

## Starting a piece of work

1. Read `WORKPLAN.md`, then the current `DESIGN_v0_...md` if one exists for
   the version in progress, then `CHANGELOG.md` and `ROADMAP.md`. Items
   that were deliberately deferred are not proposed again without saying
   so.
2. Run the test suite (`python3 -m pytest -q`) before changing anything,
   so a failure can be traced to the session that caused it.
3. If the request is open ("add X"), ask a few multiple-choice questions
   first: scope, the safe default. Do not ask what the plan already
   answers.
4. Create the branch: `feat/vX.Y.Z` for a release, `fix/vX.Y.Z` for a
   patch.
5. Build in small steps with tests. After each step, say what changed and
   what is still open.
6. When something fails in a live test, ask for the exact output or the
   log instead of guessing, and say so when the cause is not known.

## Engineering rules

1. **Never overwrite what the user did not ask to replace.** The `.sng`
   file is the app's own document: Save writes it in place, atomically,
   and a rename never replaces another song. Exports and imports never
   overwrite an existing file; they take the next free name (`name_v2`).
   The only exception is a file the user has just agreed to replace in a
   system save dialog.
2. **One writer.** Every TXT and PDF export goes through
   `export.export_to_file()` (or `prepare_export()` for a download), so the
   bytes, the check and the way files are written live in one place.
3. **Deterministic output.** The same song gives the same bytes. The only
   variables are the footer date and the version, which the golden-file
   tests fix.
4. **Validate before export.** `export.check_document()`, the same check
   as `cli.py lint`, runs before every export: errors block it, warnings
   are shown.
5. **The printed page is the reference.** Output is judged against the
   handwritten chart it replaces.
6. **Pure modules first, UI last.** `model.py` → `grammar.py` →
   `transpose.py` → `render.py` → `export.py` → front ends. Logic with no
   Tkinter in it is fully testable without a display.
7. **Fix the shared root cause** (one `ENHARMONIC` table) instead of
   patching each symptom.
8. **Dependencies:** the Python standard library only. `pytest` is
   dev-only and `pywebview` is optional (`requirements-optional.txt`).
9. **The version lives in `APP_VERSION` in `constants.py`** and nowhere
   else. Filenames are not versioned.

## Git

- One branch per release. Merge when the tests are green.
- Conventional Commits: `feat:`, `fix:`, `test:`, `docs:`, `chore:`,
  `refactor:`. One logical change per commit.
- Commit messages end with the two trailer lines the session gives
  (`Co-Authored-By:` and `Claude-Session:`).
- Do not push the feature branch. Merge locally and push `main` and the
  tag together. Pushing `main` also publishes `wiki/` through
  `.github/workflows/publish-wiki.yml`.

  ```
  git status --short                      # must print nothing
  git checkout main && git merge --ff-only fix/vX.Y.Z && git tag vX.Y.Z && git push origin main vX.Y.Z
  ```

- Before every commit, check `git status` and `git diff --cached --stat`
  for anything from the working process: `.claude/`, `Claude outputs/`,
  `Old - do not commit/`, scratch files, real song files (`*.sng`), real
  names or lyrics, keys and tokens, logs, personal paths. Never
  `git add -A` blindly.
- A tag that has not been pushed can be moved (`git tag -f`) when a fix
  joins the same release. Once pushed, make a new patch version.
- If git leaves `*.lock` or `tmp_obj_*` files behind, remove the stale
  ones, never a lock a running git may still hold.

## Testing

- Tests are written with the change, and the full suite passes before a
  commit (`python3 -m pytest -q`).
- Test the app the way it is used: the CLI through a real subprocess, the
  web front end through its real API routes, and whole songs through
  `export.py`.
- Test data uses invented band, song and lyric names. The author's songs
  stay out of the repository.
- A bug found in live use gets a test that reproduces it.
- Golden files (`tests/golden/`) record approved output. Change them only
  on purpose, with `python3 tools/make_golden.py`, and look at the new PDFs
  before committing them.
- Release notes say which systems were tried live and which were only
  built and unit-tested. "Untested" stays visible until it is not.

## Documentation

| Document | Purpose |
|---|---|
| `README.md` | What the app does and how to run it. Features, not fixes, and no "what's new". |
| `CHANGELOG.md` | Every change, under `[Unreleased]` while work is in progress, with a short "why" paragraph at the top of each version. Its section for a version is also the GitHub release body. |
| `wiki/` | The long-form pages (chart-line syntax, printing and layout, CLI and web use, architecture). New or changed behaviour updates its page in the same release. |
| `WORKPLAN.md` | The living plan: where we are, what is left, standing decisions, manual checks. |
| `ROADMAP.md` | The public wish list and the released-versions table. |
| `DESIGN_v0_...md` | The specification for a larger version, written before it is built. |

**`wiki/` in this repository is the wiki.** GitHub serves wikis from a
separate git repository. Instead of a second clone, `wiki/` is edited with
the change it describes and the workflow copies it to the wiki when it
lands on `main`. The sync is one-way and total: a page deleted from
`wiki/` is deleted from the wiki, and a page created in the wiki's web UI
is removed on the next run. If the wiki has drifted, fix `wiki/` and
re-run the workflow from the Actions tab.

**Writing style.** All prose reads as the author would write it:

- No em dashes or en dashes as connectors; use a period, comma, colon or
  parentheses.
- No "not X but Y" contrasts, no one-line closers that repeat the
  paragraph, no staged openers.
- No marketing words (powerful, seamless, robust, crucial) and no stacked
  hedges.
- Headings in sentence case. Bold only for button and menu names.
- Facts do not change when the wording does. Anything new in a text comes
  from the changelog or the code.

## Releasing

1. Bump `APP_VERSION` in `constants.py`.
2. Move `[Unreleased]` to `[X.Y.Z] — date` in `CHANGELOG.md`, with the
   "why" paragraph, and add the row to the table in `ROADMAP.md`.
3. Run the full test suite.
4. Merge, tag and push as above.
5. Create the GitHub release from the tag, with that version's
   `CHANGELOG.md` section as the body, and mark it *Latest*.
6. Small fixes are patch versions (0.26.3, 0.26.4).

## Privacy and safety

- The public repository carries no real names, songs or lyrics.
  Commits use the project author name.
- Secrets never go into the chat. If one is pasted, say so at once and
  advise creating a new one.
- Anything that talks to the network is optional and disclosed, and sends
  nothing about the user. `webserver.py` has no authentication and is
  meant for `localhost` or a trusted home network.

## Keeping the project in step

- After changing `WORKPLAN.md`, copy it to the Claude project
  (`claude/WORKPLAN.md`). The project is what a new session reads first.
- A multi-step job starts with a task list ending in a verification step.
- The codebase map under `.claude/` is optional and local only. It helps a
  fresh session, but it drifts and is not a source of truth.
