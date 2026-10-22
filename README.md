# pystrat

Stratagem picker for Helldivers 2, inspired by hellbuddy. Hotkey-driven launcher with a compact GUI, presets, and icon picker.

## Download
Grab the latest `pystrat.exe` (or the zip) from the [GitHub Releases](https://github.com/ajurna/pystrat/releases) page and run it. No Python install is needed.

Requires Windows 10/11.

## Features
- Global numpad hotkeys (NumPad 1-9, a 3x3 grid of slots) that work with the window focused or not.
- Clickable icon picker with categories and search.
- Preset loadouts with save/overwrite/delete.
- Adjustable key delay (default 40 ms) and input mode (WASD vs arrow keys).
- Cooldown tracking on each slot after activation, with optional ship module reductions.

pystrat sends the stratagem sequence as synthetic keystrokes (`SendInput`). Check the game's rules on input automation before using it.

## Using it
1. Click a slot's icon or name to open the picker, then search or browse by category and click a stratagem to assign it.
2. Press the matching numpad key (NumPad 7/8/9 is the top row, 4/5/6 the middle, 1/2/3 the bottom) to send that stratagem's sequence to the game.
3. Set **Key Delay** higher if the game misses inputs, and switch **Input Keys** between WASD and Arrows to match your in-game stratagem controls.
4. Use the **Preset** box to save your current loadout, save it under a new name, or delete it.

## Ship modules
Click **Ship Modules...** and tick the upgrades you have unlocked. Active modules shorten the cooldown shown for the stratagem types they affect. Reductions stack, up to a maximum of 95%.

## Your data
Settings, presets, keybinds, input mode, and ship modules are saved to:
```
%APPDATA%\pystrat\user_data.json
```
Delete this file to reset everything.

## Building from source
See [BUILD.md](BUILD.md) for running from source, building the exe, releasing, and adding stratagems.
