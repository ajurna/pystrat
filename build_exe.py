from __future__ import annotations

import subprocess
import sys


def build() -> None:
    args = [
        sys.executable,
        "-m",
        "PyInstaller",
        "--onefile",
        "--windowed",
        "--name",
        "pystrat",
        "--icon",
        "app.ico",
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
    subprocess.run(args, check=True)


def main() -> None:
    try:
        build()
    except subprocess.CalledProcessError as e:
        raise SystemExit(e.returncode)


if __name__ == "__main__":
    main()
