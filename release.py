from __future__ import annotations

import subprocess
import zipfile
from pathlib import Path

import build_exe


def main() -> None:
    root = Path(__file__).parent

    result = subprocess.run(
        ["uv", "version", "--short"], capture_output=True, text=True, check=True, cwd=root
    )
    version = result.stdout.strip()
    if not version:
        raise SystemExit("Could not read version from pyproject.toml")

    tag = f"v{version}"

    # Check before the 26 MB build: the tag must point at the tree that was built.
    dirty = subprocess.run(
        ["git", "status", "--porcelain"], cwd=root, capture_output=True, text=True, check=True
    )
    if dirty.stdout.strip():
        raise SystemExit("Working tree is dirty - commit before releasing.")
    if subprocess.run(["git", "rev-parse", "-q", "--verify", tag], cwd=root, check=False).returncode == 0:
        raise SystemExit(f"Tag {tag} already exists")

    release_notes = (root / "RELEASE.md").read_text()
    build_exe.clean()
    build_exe.build()

    exe_path = root / "dist" / "pystrat.exe"
    if not exe_path.exists():
        raise SystemExit(f"Build output not found: {exe_path}")

    zip_path = root / "dist" / f"pystrat-{version}.zip"
    zip_path.unlink(missing_ok=True)
    with zipfile.ZipFile(zip_path, "w", zipfile.ZIP_DEFLATED) as zf:
        zf.write(exe_path, exe_path.name)

    # Both artifacts: the exe, and a zip of it for people whose AV blocks a bare .exe.
    subprocess.run(["git", "tag", tag, "-m", tag], check=True, cwd=root)
    # Push before the release so the remote tag points at the tree that was built, not at GitHub's HEAD.
    subprocess.run(["git", "push", "origin", tag], check=True, cwd=root)
    subprocess.run(
        ["gh", "release", "create", tag, str(exe_path), str(zip_path), "-t", tag, "-n", release_notes],
        check=True,
        cwd=root,
    )


if __name__ == "__main__":
    main()
