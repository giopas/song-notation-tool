# Contributing to Song Notation Tool

Thanks for taking a look at this project.

> This is a personal project built step by step with AI assistance
> ("vibecoded"). It works, but it has rough edges, and contributions and
> feedback are welcome.

## Reporting bugs

Please open a bug with the **Bug report** issue template, and try to
include:

- your operating system and Python version (`python3 --version`);
- which front end you used (desktop app, native window, browser tab or
  command line);
- what you were doing when the bug happened;
- what you expected, and what happened instead;
- any error messages from the terminal.

## Suggesting features

Use the **Feature request** issue template. This is a personal tool, so not
every feature will fit, but every idea gets read.

## Code

There is no formal pull request process yet. To contribute code:

1. Fork the repository.
2. Make your changes on a branch with a clear name.
3. Open a pull request with a short description of what you changed and
   why.

Please keep each pull request to one feature or one fix; that makes it much
easier to review. [DEVELOPMENT.md](DEVELOPMENT.md) describes how the
project itself is worked on: tests, commits and releases.

## Style

- Python standard library only, with no external dependencies to run the
  app (`pytest` for the tests and `pywebview` for the native window are
  optional).
- Tkinter for the desktop app; plain HTML, CSS and JavaScript for the
  browser front end.
- Python 3.8 or later on macOS, Windows and Linux.
- Comments in the code are welcome. It was written quickly and can always
  be clearer.
