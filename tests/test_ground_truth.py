"""tools/ground_truth.py against small throwaway repositories.

Each check is also run against a known-bad case, so a check that can no
longer go red would show up here. The script runs as a real subprocess,
the way a session runs it.
"""

import os
import shutil
import subprocess
import sys
from pathlib import Path

import pytest

SCRIPT = Path(__file__).resolve().parent.parent / "tools" / "ground_truth.py"

pytestmark = pytest.mark.skipif(shutil.which("git") is None,
                                reason="git is not installed")

GIT_ENV = dict(os.environ,
               GIT_AUTHOR_NAME="Test", GIT_AUTHOR_EMAIL="test@example.com",
               GIT_COMMITTER_NAME="Test", GIT_COMMITTER_EMAIL="test@example.com",
               GIT_CONFIG_GLOBAL=os.devnull, GIT_CONFIG_NOSYSTEM="1")


def git(repo, *args):
    subprocess.run(["git", "-C", str(repo), *args], check=True,
                   capture_output=True, text=True, env=GIT_ENV)


def write(repo, name, text):
    (Path(repo) / name).write_text(text, encoding="utf-8")


def commit_all(repo, message):
    git(repo, "add", "-A")
    git(repo, "commit", "-q", "-m", message)


def run(repo, *extra):
    proc = subprocess.run(
        [sys.executable, str(SCRIPT), "--repo", str(repo), "--no-github", *extra],
        capture_output=True, text=True, env=GIT_ENV, timeout=120)
    return proc.returncode, proc.stdout


def marked(out):
    return [line for line in out.splitlines() if line.startswith("  ! ")]


@pytest.fixture
def repo(tmp_path):
    """A released 1.0.0: main and the tag pushed to a local 'origin'."""
    work = tmp_path / "work"
    work.mkdir()
    git(work, "init", "-q")
    git(work, "checkout", "-q", "-b", "main")
    write(work, "constants.py", 'APP_VERSION = "1.0.0"\n')
    write(work, "CHANGELOG.md", "# Changelog\n\n## [1.0.0] — 2026-01-01\n")
    write(work, "WORKPLAN.md", "# Work plan\n\n- Released: v1.0.0\n")
    commit_all(work, "chore(release): 1.0.0")
    git(work, "tag", "v1.0.0")
    git(tmp_path, "clone", "-q", "--bare", str(work), "origin.git")
    git(work, "remote", "add", "origin", str(tmp_path / "origin.git"))
    git(work, "fetch", "-q", "origin")
    git(work, "branch", "-q", "--set-upstream-to=origin/main", "main")
    return work


def test_clean_repo_is_complete_with_nothing_marked(repo):
    code, out = run(repo)
    assert code == 0, out
    assert out.rstrip().splitlines()[-1].startswith("GROUND TRUTH: complete")
    assert "UNKNOWN" not in out
    assert marked(out) == [], out
    assert "tag v1.0.0 is pushed" in out
    assert "SKIPPED: release check" in out


def test_uncommitted_change_is_marked(repo):
    write(repo, "notes.txt", "scratch\n")
    code, out = run(repo)
    assert code == 0
    assert any("uncommitted" in line for line in marked(out))
    assert "?? notes.txt" in out


def test_merged_branch_is_offered_for_deletion(repo):
    git(repo, "branch", "feat/v1.0.0")
    _, out = run(repo)
    assert any("feat/v1.0.0: merged into main" in line for line in marked(out))


def test_current_branch_without_commits_is_not_offered(repo):
    git(repo, "checkout", "-q", "-b", "chore/new")
    _, out = run(repo)
    assert "chore/new: current branch" in out
    assert not any("chore/new" in line for line in marked(out))


def test_unmerged_local_branch_is_listed(repo):
    git(repo, "checkout", "-q", "-b", "feat/v1.1.0")
    write(repo, "new.py", "x = 1\n")
    commit_all(repo, "feat: new")
    git(repo, "checkout", "-q", "main")
    _, out = run(repo)
    assert "feat/v1.1.0: 1 commit(s) not on main, local only" in out


def test_unpushed_commit_and_tag_are_marked(repo):
    write(repo, "constants.py", 'APP_VERSION = "1.0.1"\n')
    write(repo, "CHANGELOG.md", "## [1.0.1] — 2026-01-02\n")
    write(repo, "WORKPLAN.md", "- Released: v1.0.1\n")
    commit_all(repo, "chore(release): 1.0.1")
    git(repo, "tag", "v1.0.1")
    _, out = run(repo)
    lines = marked(out)
    assert any("1 commit(s) on main not pushed" in line for line in lines)
    assert any("tag v1.0.1 is not pushed" in line for line in lines)


def test_version_behind_tag_is_marked(repo):
    write(repo, "other.py", "y = 2\n")
    commit_all(repo, "fix: other")
    git(repo, "tag", "v1.1.0")
    _, out = run(repo, "--no-fetch")
    assert any("APP_VERSION 1.0.0 is behind the tag v1.1.0" in line
               for line in marked(out))


def test_version_in_progress_is_not_marked(repo):
    write(repo, "constants.py", 'APP_VERSION = "1.1.0"\n')
    commit_all(repo, "chore: start 1.1.0")
    _, out = run(repo, "--no-fetch")
    assert "version 1.1.0 is in progress" in out


def test_stale_workplan_is_marked(repo):
    write(repo, "constants.py", 'APP_VERSION = "1.0.1"\n')
    write(repo, "CHANGELOG.md", "## [1.0.1] — 2026-01-02\n")
    commit_all(repo, "chore(release): 1.0.1")
    git(repo, "tag", "v1.0.1")
    _, out = run(repo, "--no-fetch")
    lines = marked(out)
    assert any("says Released: v1.0.0, the latest tag is v1.0.1" in line
               for line in lines)
    assert any("has not changed since before v1.0.1" in line for line in lines)


def test_workplan_without_released_line_is_marked(repo):
    write(repo, "WORKPLAN.md", "# Work plan\n")
    commit_all(repo, "docs: plan")
    _, out = run(repo, "--no-fetch")
    assert any("no 'Released: vX.Y.Z' line" in line for line in marked(out))


def test_leftover_lock_file_is_marked(repo):
    (repo / ".git" / "refs" / "heads" / "old.lock").write_text("")
    _, out = run(repo, "--no-fetch")
    assert any("leftover git files" in line for line in marked(out))
    assert "old.lock" in out


def test_pushed_feature_branch_is_marked(repo, tmp_path):
    git(repo, "push", "-q", "origin", "main:feat/v1.0.0")
    _, out = run(repo)
    assert any("origin/feat/v1.0.0 is on GitHub" in line for line in marked(out))


def test_failed_fetch_is_unknown_not_a_pass(repo, tmp_path):
    shutil.rmtree(tmp_path / "origin.git")
    code, out = run(repo)
    assert code == 2
    assert "UNKNOWN: git fetch failed" in out
    assert out.rstrip().splitlines()[-1].startswith("GROUND TRUTH: INCOMPLETE")


def test_not_a_repository_is_unknown(tmp_path):
    code, out = run(tmp_path)
    assert code == 2
    assert "is not a git checkout" in out


def test_no_fetch_names_what_it_skipped(repo):
    code, out = run(repo, "--no-fetch")
    assert code == 0
    assert "SKIPPED: git fetch" in out
    assert "check(s) skipped" in out.rstrip().splitlines()[-1]
