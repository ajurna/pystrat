from __future__ import annotations

import subprocess
import sys
import zipfile
from pathlib import Path

import build_exe


def main() -> None:
    root = Path(__file__).parent

    result = subprocess.run(["uv", "version", "--short"], capture_output=True, text=True, check=True)
    version = result.stdout.strip()
    if not version:
        raise SystemExit("Could not read version from pyproject.toml")

    tag = f"v{version}"

    build_exe.build()

    exe_path = root / "dist" / "pystrat.exe"
    if not exe_path.exists():
        raise SystemExit(f"Build output not found: {exe_path}")

    zip_path = root / "dist" / f"pystrat-{version}.zip"
    zip_path.unlink(missing_ok=True)
    with zipfile.ZipFile(zip_path, "w", zipfile.ZIP_DEFLATED) as zf:
        zf.write(exe_path, exe_path.name)

    release_notes = (root / "RELEASE.md").read_text()

    subprocess.run(["git", "tag", tag], check=True)
    subprocess.run(["git", "push", "origin", tag], check=True)
    subprocess.run(
        ["gh", "release", "create", tag, str(exe_path), str(zip_path), "-t", tag, "-n", release_notes],
        check=True,
    )


if __name__ == "__main__":
    main()
