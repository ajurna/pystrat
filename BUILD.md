# Building pystrat

Developer notes for running, building, and releasing pystrat. For end-user information see [README.md](README.md).

## Requirements
- Windows 10/11
- Python 3.13+
- [uv](https://docs.astral.sh/uv/) (recommended)

Runtime dependencies are listed in `pyproject.toml`.

## Run (dev)
From the project root:
```
uv sync
uv run python main.py
```

## Build (exe)
PyInstaller lives in the `build` dependency group, so install it first:
```
uv sync --group build
uv run --group build python build_exe.py
```

The output is `dist\pystrat.exe`.

`build_exe.py` runs PyInstaller under the base interpreter rather than uv's venv trampoline, which intermittently fails with `WinError 5` on Windows. It also cleans `build\` and any `*.spec` file before each build.

## Release
`release.py` builds the exe, zips it, tags `v<version>`, pushes the tag, and creates a GitHub release using the contents of `RELEASE.md` as notes. It needs `uv`, `git`, and an authenticated `gh` CLI. Bump `version` in `pyproject.toml` and update `RELEASE.md` first:
```
uv run --group build release
```

## Data files
- `stratagems.json` - list of stratagems, sequences, categories, and cooldowns.
- `StratagemIcons\` - SVG icon files. Filenames must match stratagem names.
- `app.ico` - window/taskbar icon.

`pyproject.toml` is also bundled into the exe so the app can read its version.

### Adding a stratagem
Add an entry to `stratagems.json` and drop a matching `<name>.svg` into `StratagemIcons\`:
```json
{"name": "Orbital Precision Strike", "sequence": ["D", "D", "W"], "category": "Offensive", "cooldown_seconds": 90, "type": "orbital"}
```
- `sequence` uses `W`, `A`, `S`, `D` (translated to arrow keys when that input mode is selected).
- `category` is one of `Offensive`, `Defensive`, `General`, or `Supply`, and drives the icon picker tabs.
- `type` is one of `orbital`, `eagle`, `support_weapon`, `backpack`, `sentry`, `emplacement`, `resupply`, `vehicle`, `exosuit`, or `mission`. Ship modules use it to adjust cooldowns.
- `cooldown_seconds` may be omitted or empty for stratagems with no cooldown arc.

## Notes
- The icon picker remembers scroll position during the current app session only.
