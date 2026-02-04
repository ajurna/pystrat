# pystrat

Hotkey-driven stratagem launcher with a compact GUI, presets, and icon picker.

## Features
- Global numpad hotkeys (works with the window focused or not).
- Clickable icon picker with categories and search.
- Preset loadouts with save/overwrite/delete.
- Adjustable key delay and input mode (WASD vs arrow keys).

## Requirements
- Windows 10/11
- Python 3.14+

Runtime dependencies are listed in `pyproject.toml`:
- `svglib`
- `reportlab`
- `pillow`

## Run (dev)
From the project root:
```
python main.py
```

## Build (exe)
The build script is `build_exe.py`. Run it via uv or python:
```
uv run -- python build_exe.py
```
or
```
python build_exe.py
```

The output will be in `dist\main.exe`.

## Data files
- `stratagems.json` - list of stratagems, sequences, and categories.
- `StratagemIcons\` - SVG icon files. Filenames must match stratagem names.
- `app.ico` - window/taskbar icon.

## User data
Settings and presets are saved to:
```
%APPDATA%\pystrat\user_data.json
```

## Notes
- The icon picker categories are driven by the `category` field in `stratagems.json`.
- The selector remembers scroll position during the current app session only.
