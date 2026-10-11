#!/usr/bin/env python3
"""Ground truth: what git, the tags and GitHub say about the project now.

    python3 tools/ground_truth.py               # full check
    python3 tools/ground_truth.py --no-fetch    # no network at all
    python3 tools/ground_truth.py --no-github   # skip the release check

WORKPLAN.md is what an earlier session wrote down. This script prints the
facts it should agree with: the working tree, `main` against GitHub, the
version in constants.py against the tags, the branches, the GitHub release
of the latest tag, and stale git lock files. A session runs it first and
corrects WORKPLAN.md where the two disagree.

It changes nothing, apart from `git fetch --prune origin`. Lines marked
`!` need a look. A probe that could not run is printed as UNKNOWN and is
never read as a pass. The last line says whether the report is complete:

    exit 0  GROUND TRUTH: complete (skipped checks are named)
    exit 2  GROUND TRUTH: INCOMPLETE, one or more probes failed

Python standard library only. The release check reads the public GitHub
API and sends nothing but the repository name.
"""

import argparse
import json
import os
import re
import subprocess
import sys
import urllib.error
import urllib.request
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
MAIN = "main"
TAG_RE = re.compile(r"^v(\d+)\.(\d+)\.(\d+)$")
# The first "Released: vX.Y.Z" in WORKPLAN.md, bold or not.
RELEASED_RE = re.compile(r"Released:\**\s*v(\d+\.\d+\.\d+)")


class Report:
    """Prints the report and counts what needs attention."""

    def __init__(self, out=sys.stdout):
        self.out = out
        self.unknown = 0
        self.warnings = 0
        self.skipped = 0

    def section(self, title):
        print(f"\n== {title}", file=self.out)

    def line(self, text):
        print(f"  {text}", file=self.out)

    def warn(self, text):
        self.warnings += 1
        print(f"  ! {text}", file=self.out)

    def fail(self, text):
        self.unknown += 1
        print(f"  UNKNOWN: {text}", file=self.out)

    def skip(self, text):
        self.skipped += 1
        print(f"  SKIPPED: {text}", file=self.out)

    def finish(self):
        print("", file=self.out)
        if self.unknown:
            print(f"GROUND TRUTH: INCOMPLETE, {self.unknown} probe(s) failed; "
                  "read those sections as UNKNOWN", file=self.out)
            return 2
        tail = f", {self.warnings} item(s) marked !" if self.warnings else ""
        if self.skipped:
            tail += f", {self.skipped} check(s) skipped"
        print(f"GROUND TRUTH: complete{tail}", file=self.out)
        return 0


def git(repo, *args, timeout=30):
    """Run git in repo. Returns (exit code, stdout, stderr), never raises."""
    env = dict(os.environ, GIT_TERMINAL_PROMPT="0", LC_ALL="C")
    try:
        proc = subprocess.run(["git", "-C", str(repo), *args],
                              capture_output=True, text=True,
                              timeout=timeout, env=env)
    except FileNotFoundError:
        return 127, "", "git is not installed"
    except subprocess.TimeoutExpired:
        return 124, "", f"timed out after {timeout} s"
    return proc.returncode, proc.stdout.strip(), proc.stderr.strip()


def first_line(text):
    return text.splitlines()[0] if text else "(no message)"


def version_key(tag):
    m = TAG_RE.match(tag)
    return tuple(int(x) for x in m.groups()) if m else None


def read_app_version(repo):
    try:
        text = (Path(repo) / "constants.py").read_text(encoding="utf-8")
    except OSError:
        return None
    m = re.search(r'^APP_VERSION\s*=\s*"([^"]+)"', text, re.M)
    return m.group(1) if m else None


def changelog_has(repo, version):
    try:
        text = (Path(repo) / "CHANGELOG.md").read_text(encoding="utf-8")
    except OSError:
        return None
    return re.search(rf"^##\s*\[{re.escape(version)}\]", text, re.M) is not None


def github_slug(url):
    m = re.search(r"github\.com[:/]([^/]+)/([^/]+?)(?:\.git)?/?$", url or "")
    return f"{m.group(1)}/{m.group(2)}" if m else None


def github_get(path, timeout=15):
    """GET from the public GitHub API. Returns (status, data or None)."""
    req = urllib.request.Request(
        f"https://api.github.com/{path}",
        headers={"Accept": "application/vnd.github+json",
                 "User-Agent": "song-notation-tool-ground-truth"})
    try:
        with urllib.request.urlopen(req, timeout=timeout) as resp:
            return resp.status, json.loads(resp.read().decode("utf-8"))
    except urllib.error.HTTPError as err:
        return err.code, None


def is_ancestor(repo, a, b):
    code, _, _ = git(repo, "merge-base", "--is-ancestor", a, b)
    return code == 0


def check(repo, report, fetch=True, use_github=True):
    repo = Path(repo)

    report.section("Checkout")
    code, out, err = git(repo, "rev-parse", "--show-toplevel")
    if code != 0:
        report.fail(f"{repo} is not a git checkout ({err or 'no output'})")
        return
    report.line(f"repository: {out}")
    code, _, _ = git(repo, "rev-parse", "--verify", "--quiet", MAIN)
    if code != 0:
        report.fail(f"there is no local branch {MAIN}")
        return

    code, origin_url, _ = git(repo, "remote", "get-url", "origin")
    has_origin = code == 0 and bool(origin_url)
    remote_tags = None
    if not has_origin:
        report.warn("no remote 'origin': nothing below is compared with GitHub")
    elif not fetch:
        report.skip("git fetch and the remote tag list (--no-fetch); "
                    "remote comparisons use the last fetch")
    else:
        code, _, err = git(repo, "fetch", "--prune", "origin", timeout=60)
        if code != 0:
            report.fail(f"git fetch failed ({first_line(err)}); remote "
                        "comparisons below use the last successful fetch")
        code, out, err = git(repo, "ls-remote", "--tags", "origin", timeout=60)
        if code != 0:
            report.fail(f"could not list the tags on origin ({first_line(err)})")
        else:
            remote_tags = {}
            for row in out.splitlines():
                sha, _, ref = row.partition("\t")
                if ref.startswith("refs/tags/") and not ref.endswith("^{}"):
                    remote_tags[ref[len("refs/tags/"):]] = sha

    report.section("Working tree")
    code, out, err = git(repo, "status", "--porcelain=v1", "--branch")
    if code != 0:
        report.fail(f"git status failed ({first_line(err)})")
    else:
        rows = out.splitlines()
        if rows and rows[0].startswith("## "):
            report.line(f"on branch: {rows[0][3:]}")
            rows = rows[1:]
        if rows:
            report.warn(f"{len(rows)} uncommitted change(s): commit them, or "
                        "name them in WORKPLAN.md")
            for row in rows[:20]:
                report.line(f"    {row}")
            if len(rows) > 20:
                report.line(f"    ... and {len(rows) - 20} more")
        else:
            report.line("clean")

    report.section(f"Branch {MAIN}")
    _, out, _ = git(repo, "log", "-1", "--format=%h %cs %s", MAIN)
    report.line(f"last commit: {out}")
    code, out, _ = git(repo, "rev-list", "--left-right", "--count",
                       f"{MAIN}...origin/{MAIN}")
    if not has_origin:
        pass
    elif code != 0:
        report.warn(f"origin/{MAIN} not found: {MAIN} has never been pushed")
    else:
        ahead, behind = (int(x) for x in out.split())
        if ahead == 0 and behind == 0:
            report.line(f"same as origin/{MAIN}")
        else:
            if ahead:
                report.warn(f"{ahead} commit(s) on {MAIN} not pushed")
            if behind:
                report.warn(f"{behind} commit(s) on origin/{MAIN} not pulled")

    report.section("Version and tags")
    version = read_app_version(repo)
    code, out, err = git(repo, "tag", "--list", "v*")
    if code != 0:
        report.fail(f"could not list tags ({first_line(err)})")
        tags = []
    else:
        tags = sorted((t for t in out.split() if version_key(t)),
                      key=version_key)
    on_main = [t for t in tags if is_ancestor(repo, t, MAIN)]
    for tag in tags:
        if tag not in on_main:
            report.warn(f"tag {tag} is not on {MAIN}")
    latest = on_main[-1] if on_main else None
    report.line(f"APP_VERSION in constants.py: {version or 'not found'}")
    report.line(f"latest tag on {MAIN}: {latest or 'none'}")
    if version is None:
        report.fail("APP_VERSION not found in constants.py")
    elif latest:
        v_key = version_key("v" + version)
        t_key = version_key(latest)
        if v_key is None:
            report.warn(f"APP_VERSION {version} is not of the form X.Y.Z")
        elif v_key == t_key:
            report.line(f"{MAIN} carries the released version {version}")
        elif v_key > t_key:
            report.line(f"version {version} is in progress (not tagged yet)")
        else:
            report.warn(f"APP_VERSION {version} is behind the tag {latest}")
    if version and changelog_has(repo, version) is False and latest == f"v{version}":
        report.warn(f"CHANGELOG.md has no section [{version}]")
    if latest:
        _, out, _ = git(repo, "rev-list", "--count", f"{latest}..{MAIN}")
        report.line(f"commits on {MAIN} since {latest}: {out}")
        if remote_tags is not None:
            _, local_sha, _ = git(repo, "rev-parse", f"refs/tags/{latest}")
            if latest not in remote_tags:
                report.warn(f"tag {latest} is not pushed")
            elif remote_tags[latest] != local_sha:
                report.warn(f"tag {latest} on origin points elsewhere than "
                            "the local one")
            else:
                report.line(f"tag {latest} is pushed")

    report.section("Branches")
    code, out, err = git(repo, "for-each-ref", "refs/heads",
                         "--format=%(refname:short)\t%(upstream:short)")
    if code != 0:
        report.fail(f"could not list branches ({first_line(err)})")
    else:
        others = [r.split("\t") for r in out.splitlines()
                  if r.split("\t")[0] != MAIN]
        if not others:
            report.line(f"only {MAIN}")
        _, current, _ = git(repo, "branch", "--show-current")
        for name, upstream in others:
            if name == current and is_ancestor(repo, name, MAIN):
                report.line(f"{name}: current branch, nothing on it yet")
                continue
            if is_ancestor(repo, name, MAIN):
                report.warn(f"{name}: merged into {MAIN}, can be deleted "
                            f"(git branch -d {name})")
                continue
            _, ahead, _ = git(repo, "rev-list", "--count", f"{MAIN}..{name}")
            where = f"tracks {upstream}" if upstream else "local only"
            report.line(f"{name}: {ahead} commit(s) not on {MAIN}, {where}")
    if has_origin:
        code, out, _ = git(repo, "for-each-ref", "refs/remotes/origin",
                           "--format=%(refname:short)")
        for ref in out.splitlines() if code == 0 else []:
            if ref in ("origin/HEAD", f"origin/{MAIN}", "origin"):
                continue
            name = ref[len("origin/"):]
            state = ("merged into " + MAIN) if is_ancestor(repo, ref, MAIN) \
                else "not merged"
            report.warn(f"{ref} is on GitHub ({state}); feature branches stay "
                        f"local. To remove it: git push origin --delete {name}")

    report.section("Git housekeeping")
    code, git_dir, _ = git(repo, "rev-parse", "--absolute-git-dir")
    if code != 0:
        report.fail("could not find the .git folder")
    else:
        stale = sorted(p for p in Path(git_dir).rglob("*")
                       if p.name.endswith(".lock") or p.name.startswith("tmp_obj_"))
        if stale:
            report.warn("leftover git files; remove them only when no git "
                        "command is running:")
            for p in stale[:10]:
                report.line(f"    {p.relative_to(Path(git_dir).parent)}")
        else:
            report.line("no leftover lock or tmp_obj files")

    report.section("WORKPLAN.md")
    try:
        plan = (repo / "WORKPLAN.md").read_text(encoding="utf-8")
    except OSError:
        plan = None
    if plan is None:
        report.warn("there is no WORKPLAN.md")
    else:
        m = RELEASED_RE.search(plan)
        if not m:
            report.warn("WORKPLAN.md has no 'Released: vX.Y.Z' line")
        elif latest and f"v{m.group(1)}" != latest:
            report.warn(f"WORKPLAN.md says Released: v{m.group(1)}, the latest "
                        f"tag is {latest}: update 'Where we are'")
        else:
            report.line(f"says Released: v{m.group(1)}")
    code, wp_commit, _ = git(repo, "log", "-1", "--format=%H", MAIN, "--",
                             "WORKPLAN.md")
    if code != 0 or not wp_commit:
        report.warn("WORKPLAN.md is not committed on main")
    else:
        _, when, _ = git(repo, "log", "-1", "--format=%cs %s", wp_commit)
        report.line(f"last change: {when}")
        if latest:
            _, tag_commit, _ = git(repo, "rev-parse", f"{latest}^{{commit}}")
            if wp_commit != tag_commit and is_ancestor(repo, wp_commit, tag_commit):
                report.warn(f"WORKPLAN.md has not changed since before {latest}: "
                            "check its 'Where we are' against this report")

    report.section("GitHub release")
    slug = github_slug(origin_url) if has_origin else None
    if not latest:
        report.line("no tag yet")
    elif not use_github or not fetch:
        report.skip("release check (--no-github or --no-fetch)")
    elif not slug:
        report.skip("origin is not a GitHub repository")
    else:
        try:
            status, data = github_get(f"repos/{slug}/releases/tags/{latest}")
            latest_status, latest_data = github_get(f"repos/{slug}/releases/latest")
        except (urllib.error.URLError, OSError, ValueError) as err:
            report.fail(f"GitHub API not reachable ({err})")
        else:
            if status == 200 and data:
                report.line(f"{latest} published {str(data.get('published_at'))[:10]}")
            elif status == 404:
                report.warn(f"{latest} has no published release (a draft, or "
                            "none). Update and restart offers only published "
                            "releases")
            else:
                report.fail(f"GitHub answered {status} for {latest}")
            if latest_status == 200 and latest_data:
                marked = latest_data.get("tag_name")
                if marked != latest and status == 200:
                    report.warn(f"GitHub marks {marked} as Latest, not {latest}")
            elif latest_status != 404:
                report.fail(f"GitHub answered {latest_status} for the latest release")


def main(argv=None):
    parser = argparse.ArgumentParser(description=__doc__.splitlines()[0])
    parser.add_argument("--repo", default=str(ROOT),
                        help="repository to check (default: this one)")
    parser.add_argument("--no-fetch", action="store_true",
                        help="no network: skip git fetch, the remote tag "
                             "list and GitHub")
    parser.add_argument("--no-github", action="store_true",
                        help="skip the GitHub release check")
    args = parser.parse_args(argv)
    report = Report()
    from datetime import datetime
    print(f"Ground truth, {datetime.now().strftime('%Y-%m-%d %H:%M')}")
    check(args.repo, report, fetch=not args.no_fetch,
          use_github=not args.no_github)
    return report.finish()


if __name__ == "__main__":
    sys.exit(main())
