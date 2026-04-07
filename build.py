"""
Build script for Reconciliator.

Usage:
    python build.py                  # Build exe with current version
    python build.py --bump patch     # Bump patch (1.0.0 -> 1.0.1) and build
    python build.py --bump minor     # Bump minor (1.0.0 -> 1.1.0) and build
    python build.py --bump major     # Bump major (1.0.0 -> 2.0.0) and build
    python build.py --bump patch --no-build   # Only bump, skip exe build
    python build.py --bump patch --release    # Bump, commit, tag, push (CI builds & releases)
"""
import argparse
import os
import re
import subprocess
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent
VERSION_FILE = ROOT / "src" / "config" / "_version.py"
VERSION_PATTERN = re.compile(r'^__version__\s*=\s*"(\d+\.\d+\.\d+)"', re.MULTILINE)


def read_version() -> str:
    text = VERSION_FILE.read_text()
    m = VERSION_PATTERN.search(text)
    if not m:
        sys.exit(f"Could not parse version from {VERSION_FILE}")
    return m.group(1)


def write_version(version: str) -> None:
    text = VERSION_FILE.read_text()
    new_text = VERSION_PATTERN.sub(f'__version__ = "{version}"', text)
    VERSION_FILE.write_text(new_text)
    print(f"  Version updated to {version} in {VERSION_FILE.relative_to(ROOT)}")


def bump(part: str) -> str:
    old = read_version()
    major, minor, patch = (int(x) for x in old.split("."))
    if part == "major":
        major, minor, patch = major + 1, 0, 0
    elif part == "minor":
        minor, patch = minor + 1, 0
    elif part == "patch":
        patch += 1
    else:
        sys.exit(f"Unknown bump part: {part!r}")
    new = f"{major}.{minor}.{patch}"
    write_version(new)
    return new


def build_exe(version: str) -> None:
    dist_name = f"Reconciliator-v{version}"
    cmd = [
        sys.executable, "-m", "PyInstaller",
        "--noconfirm",
        "--onefile",
        "--windowed",
        f"--name={dist_name}",
        "--paths=src",
        # Bundle resource files
        f"--add-data=resources/icons{os.pathsep}resources/icons",
        f"--add-data=resources/sql{os.pathsep}resources/sql",
        f"--add-data=resources/demo{os.pathsep}resources/demo",
        "src/core/main.py",
    ]
    print(f"\n  Building {dist_name} ...")
    print(f"  Command: {' '.join(cmd)}\n")
    result = subprocess.run(cmd, cwd=str(ROOT))
    if result.returncode != 0:
        sys.exit(f"PyInstaller exited with code {result.returncode}")
    print(f"\n  Build complete: dist/{dist_name}.exe")


def git_release(version: str) -> None:
    """Commit the version bump, create a git tag, and push to trigger the release workflow."""
    tag = f"v{version}"

    def run_git(*args: str) -> None:
        result = subprocess.run(["git", *args], cwd=str(ROOT))
        if result.returncode != 0:
            sys.exit(f"git {args[0]} failed (exit {result.returncode})")

    run_git("add", str(VERSION_FILE.relative_to(ROOT)))
    run_git("commit", "-m", f"release: {tag}")
    run_git("tag", "-a", tag, "-m", f"Release {tag}")
    run_git("push", "--follow-tags", "--force")
    print(f"\n  Pushed tag {tag} — GitHub Actions will build & publish the release.")


def main():
    parser = argparse.ArgumentParser(description="Bump version and build Reconciliator executable.")
    parser.add_argument("--bump", choices=["major", "minor", "patch"], help="Version part to bump before building.")
    parser.add_argument("--no-build", action="store_true", help="Only bump the version; skip the executable build.")
    parser.add_argument("--release", action="store_true", help="Commit version bump, tag, and push to trigger CI release.")
    args = parser.parse_args()

    if args.release and not args.bump:
        sys.exit("--release requires --bump (e.g. --bump patch --release)")

    if args.bump:
        version = bump(args.bump)
    else:
        version = read_version()

    print(f"\n  Reconciliator v{version}")

    if args.release:
        git_release(version)
        print("  Skipping local build — CI will handle it.")
    elif not args.no_build:
        build_exe(version)


if __name__ == "__main__":
    main()
