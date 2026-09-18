# Song Notation Tool — Codebase Map

*Last Updated: 2026-09-18*

A compact one-page song charting app for bass and guitar. Three frontends (Tkinter desktop,
browser/pywebview, headless CLI) share a pure-Python engine with zero external dependencies.
Current version: **v0.19.0**, document format version 2.

## Map Documents

| Document                                        | What it covers                                        |
|-------------------------------------------------|-------------------------------------------------------|
| [architecture.md](architecture.md)              | System overview, three-frontend design, data flow     |
| [tech-landscape.md](tech-landscape.md)          | Language, deps (zero core), source-of-truth files     |
| [directory-structure.md](directory-structure.md) | Annotated folder tree, local-only dirs                |
| [entry-points.md](entry-points.md)              | Run commands, REST API routes, CLI subcommands        |
| [modules.md](modules.md)                        | Every module: purpose, key functions, exports         |
| [patterns.md](patterns.md)                      | Design patterns, error handling, testing, style       |
| [onboarding.md](onboarding.md)                  | Quick start, common tasks, file format, GitHub link   |

## How to use this map

Read `INDEX.md` (this file) for orientation. Open a detail doc when you need specifics — the
table above tells you which one to pick. Each doc is self-contained; cross-links point to
related docs when useful.

## How to maintain this map

After structural changes (new modules, moved files, schema changes, new entry points), update
the affected doc(s) and the `Last Updated` date. Run the state writer script to refresh
`.map-state.json`. The `update-codebase-map` skill handles this automatically when invoked.
