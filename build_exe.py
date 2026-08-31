from __future__ import annotations

import os
import shutil
import subprocess
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent


def real_python() -> str:
    """Return the real base interpreter, never uv's venv trampoline.

    uv creates ``.venv\\Scripts\\python.exe`` as a lightweight trampoline. That
    trampoline intermittently fails to spawn Python child processes on Windows
    with "permission denied (os error 5)". PyInstaller spawns many isolated
    child processes during analysis, so running it through the trampoline makes
    the build crash with ``PermissionError: [WinError 5]``. Running PyInstaller
    under the real interpreter means every child it spawns is the real
    interpreter too.
    """
    return getattr(sys, "_base_executable", None) or sys.executable


def build_env() -> dict[str, str]:
    """Environment that lets the base interpreter import the venv's packages."""
    env = os.environ.copy()
    site_packages = Path(sys.prefix) / ("Lib/site-packages" if os.name == "nt" else "lib")
    existing = env.get("PYTHONPATH", "")
    env["PYTHONPATH"] = os.pathsep.join(p for p in (str(site_packages), existing) if p)
    return env


def clean() -> None:
    """Remove stale PyInstaller output so a rebuild never reuses old analysis."""
    targets = [ROOT / "build", ROOT / "dist" / "pystrat.exe", *ROOT.glob("*.spec")]
    for path in targets:
        if path.is_dir():
            shutil.rmtree(path, ignore_errors=True)
        elif path.exists():
            path.unlink()


def build() -> None:
    args = [
        real_python(),
        "-m",
        "PyInstaller",
        "--noconfirm",
        "--clean",
        "--onefile",
        "--windowed",
        "--name",
        "pystrat",
        "--icon",
        "app.ico",
        "--hidden-import",
        "tkinter",
        "--hidden-import",
        "_tkinter",
        "--add-data",
        "pyproject.toml;.",
        "--add-data",
        "stratagems.json;.",
        "--add-data",
        "StratagemIcons;StratagemIcons",
        "--add-data",
        "app.ico;.",
        "main.py",
    ]
    subprocess.run(args, check=True, cwd=ROOT, env=build_env())


def main() -> None:
    try:
        clean()
        build()
    except subprocess.CalledProcessError as e:
        raise SystemExit(e.returncode)


if __name__ == "__main__":
    main()
