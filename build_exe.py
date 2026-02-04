from __future__ import annotations

import subprocess
import sys


def main() -> None:
    args = [
        sys.executable,
        "-m",
        "PyInstaller",
        "--onefile",
        "--windowed",
        "--icon",
        "app.ico",
        "--add-data",
        "stratagems.json;.",
        "--add-data",
        "StratagemIcons;StratagemIcons",
        "--add-data",
        "app.ico;.",
        "main.py",
    ]
    raise SystemExit(subprocess.call(args))


if __name__ == "__main__":
    main()
