from __future__ import annotations

import json
import os
import threading
import time
from dataclasses import dataclass
from pathlib import Path
from typing import Callable, Dict, List, Optional, Tuple

import tkinter as tk
from tkinter import ttk

SVG_ERROR = None
try:
    import cairosvg  # type: ignore
    from PIL import Image, ImageTk  # type: ignore

    SVG_AVAILABLE = True
except Exception as exc:
    SVG_AVAILABLE = False
    SVG_ERROR = str(exc)


APP_TITLE = "Stratagem Hotkeys"
BASE_DIR = Path(__file__).resolve().parent
DATA_FILE = BASE_DIR / "user_data.json"
STRATAGEMS_FILE = BASE_DIR / "stratagems.json"
ICON_DIR = BASE_DIR / "StratagemIcons"
ICON_CACHE_DIR = BASE_DIR / ".icon_cache"

DARK_BG = "#131316"
CARD_BG = "#1c1c22"
TEXT_FG = "#f1f1f3"
MUTED_FG = "#a2a2ad"

KEY_VK = {
    "W": 0x57,
    "A": 0x41,
    "S": 0x53,
    "D": 0x44,
}

WM_HOTKEY = 0x0312
MOD_NOREPEAT = 0x4000


@dataclass
class Stratagem:
    name: str
    sequence: List[str]


def load_stratagems() -> List[Stratagem]:
    with STRATAGEMS_FILE.open("r", encoding="utf-8") as handle:
        raw = json.load(handle)
    return [Stratagem(entry["name"], entry["sequence"]) for entry in raw]


def load_user_data() -> Dict[str, List]:
    if not DATA_FILE.exists():
        return {"equipped_stratagems": [], "keybinds": []}
    with DATA_FILE.open("r", encoding="utf-8") as handle:
        return json.load(handle)


def save_user_data(equipped: List[str], keybinds: List[Dict[str, str]]) -> None:
    payload = {"equipped_stratagems": equipped, "keybinds": keybinds}
    with DATA_FILE.open("w", encoding="utf-8") as handle:
        json.dump(payload, handle, indent=4)


def safe_cache_name(name: str, size: int) -> str:
    safe = "".join(ch if ch.isalnum() else "_" for ch in name)
    return f"{safe}_{size}.png"


def render_svg_to_png(svg_path: Path, size: int) -> Optional[Path]:
    if not SVG_AVAILABLE:
        return None
    ICON_CACHE_DIR.mkdir(exist_ok=True)
    cache_path = ICON_CACHE_DIR / safe_cache_name(svg_path.stem, size)
    if not cache_path.exists() or cache_path.stat().st_mtime < svg_path.stat().st_mtime:
        cairosvg.svg2png(
            url=str(svg_path),
            write_to=str(cache_path),
            output_width=size,
            output_height=size,
        )
    return cache_path


def parse_key_code(code: str) -> int:
    return int(code, 16)


class HotkeyManager:
    def __init__(self) -> None:
        if os.name != "nt":
            raise RuntimeError("Hotkeys require Windows.")
        import ctypes
        from ctypes import wintypes

        self.ctypes = ctypes
        self.wintypes = wintypes
        self.user32 = ctypes.windll.user32
        self.handlers: Dict[int, Callable[[], None]] = {}

        class MSG(ctypes.Structure):
            _fields_ = [
                ("hwnd", wintypes.HWND),
                ("message", wintypes.UINT),
                ("wParam", wintypes.WPARAM),
                ("lParam", wintypes.LPARAM),
                ("time", wintypes.DWORD),
                ("pt", wintypes.POINT),
            ]

        self.MSG = MSG

    def register(self, hotkey_id: int, vk: int, handler: Callable[[], None]) -> None:
        if not self.user32.RegisterHotKey(None, hotkey_id, MOD_NOREPEAT, vk):
            raise RuntimeError(f"Failed to register hotkey {hotkey_id}")
        self.handlers[hotkey_id] = handler

    def unregister_all(self) -> None:
        for hotkey_id in list(self.handlers.keys()):
            self.user32.UnregisterHotKey(None, hotkey_id)
        self.handlers.clear()

    def pump(self, max_messages: int = 25) -> None:
        msg = self.MSG()
        for _ in range(max_messages):
            if not self.user32.PeekMessageW(self.ctypes.byref(msg), None, 0, 0, 1):
                break
            if msg.message == WM_HOTKEY:
                handler = self.handlers.get(msg.wParam)
                if handler:
                    handler()


class StratagemApp:
    def __init__(self, root: tk.Tk) -> None:
        self.root = root
        self.root.title(APP_TITLE)
        self.root.configure(bg=DARK_BG)
        self.root.geometry("820x540")

        self.stratagems = load_stratagems()
        self.stratagem_map = {item.name: item for item in self.stratagems}
        self.stratagem_names = [item.name for item in self.stratagems]

        self.user_data = load_user_data()
        self.keybinds = self.user_data.get("keybinds", [])
        self.equipped = self.user_data.get("equipped_stratagems", [])

        desired_keybinds = [
            {"key_code": "0x67", "letter": "NumPad7"},
            {"key_code": "0x68", "letter": "NumPad8"},
            {"key_code": "0x69", "letter": "NumPad9"},
            {"key_code": "0x64", "letter": "NumPad4"},
            {"key_code": "0x65", "letter": "NumPad5"},
            {"key_code": "0x66", "letter": "NumPad6"},
            {"key_code": "0x61", "letter": "NumPad1"},
            {"key_code": "0x62", "letter": "NumPad2"},
            {"key_code": "0x63", "letter": "NumPad3"},
        ]

        if not self.keybinds:
            self.keybinds = list(desired_keybinds)
        else:
            existing = {entry.get("letter") for entry in self.keybinds}
            missing = [entry for entry in desired_keybinds if entry["letter"] not in existing]
            if missing:
                self.keybinds.extend(missing)
            order_map = {entry["letter"]: idx for idx, entry in enumerate(desired_keybinds)}
            self.keybinds.sort(key=lambda entry: order_map.get(entry.get("letter"), 999))

        for name in self.equipped:
            if name and name not in self.stratagem_names:
                self.stratagem_names.append(name)

        if len(self.equipped) < len(self.keybinds):
            default_fill = self.stratagem_names[: len(self.keybinds) - len(self.equipped)]
            self.equipped.extend(default_fill)
        self.equipped = self.equipped[: len(self.keybinds)]
        save_user_data(self.equipped, self.keybinds)

        self.icon_cache: Dict[Tuple[str, int], Optional["ImageTk.PhotoImage"]] = {}
        self.icon_jobs: set[Tuple[str, int]] = set()
        self.slot_vars: List[tk.StringVar] = []
        self.sequence_labels: List[tk.Label] = []
        self.icon_labels: List[tk.Label] = []

        self.status_var = tk.StringVar(value="Ready")
        self.last_icon_error: Optional[str] = None

        self.hotkeys: Optional[HotkeyManager] = None
        self.build_ui()
        self.register_hotkeys()

    def build_ui(self) -> None:
        header = tk.Label(
            self.root,
            text="Stratagem Hotkeys",
            bg=DARK_BG,
            fg=TEXT_FG,
            font=("Segoe UI", 18, "bold"),
        )
        header.pack(pady=(18, 6))

        subtitle = tk.Label(
            self.root,
            text="Press a numpad key to activate the assigned stratagem.",
            bg=DARK_BG,
            fg=MUTED_FG,
            font=("Segoe UI", 10),
        )
        subtitle.pack(pady=(0, 12))

        grid_frame = tk.Frame(self.root, bg=DARK_BG)
        grid_frame.pack(fill="both", expand=True, padx=20, pady=10)

        columns = 3
        rows = 3
        for col in range(columns):
            grid_frame.grid_columnconfigure(col, weight=1)
        for row in range(rows):
            grid_frame.grid_rowconfigure(row, weight=1)

        for index, keybind in enumerate(self.keybinds):
            row = index // columns
            col = index % columns
            card = tk.Frame(grid_frame, bg=CARD_BG, padx=10, pady=10)
            card.grid(row=row, column=col, sticky="nsew", padx=8, pady=8)

            key_label = tk.Label(
                card,
                text=keybind["letter"],
                bg=CARD_BG,
                fg=TEXT_FG,
                font=("Segoe UI", 11, "bold"),
            )
            key_label.grid(row=0, column=0, sticky="w")

            icon_label = tk.Label(card, bg="#0f0f12", width=64, height=64, anchor="center")
            icon_label.grid(row=1, column=0, rowspan=2, padx=(0, 12), pady=6)

            strat_var = tk.StringVar(value=self.equipped[index])
            self.slot_vars.append(strat_var)

            combo = ttk.Combobox(
                card,
                textvariable=strat_var,
                values=self.stratagem_names,
                state="readonly",
                width=30,
            )
            combo.grid(row=1, column=1, sticky="ew")
            combo.bind("<<ComboboxSelected>>", lambda _e, i=index: self.on_stratagem_change(i))

            seq_label = tk.Label(
                card,
                text=self.sequence_for(self.equipped[index]),
                bg=CARD_BG,
                fg=MUTED_FG,
                font=("Consolas", 11),
            )
            seq_label.grid(row=2, column=1, sticky="w")

            self.sequence_labels.append(seq_label)
            self.icon_labels.append(icon_label)

            self.update_icon(index)

        status_frame = tk.Frame(self.root, bg=CARD_BG, height=28)
        status_frame.pack(side="bottom", fill="x")
        status_frame.pack_propagate(False)
        status = tk.Label(
            status_frame,
            textvariable=self.status_var,
            bg=CARD_BG,
            fg=TEXT_FG,
            font=("Segoe UI", 9),
            anchor="w",
            padx=12,
        )
        status.pack(fill="both", expand=True)

        if not SVG_AVAILABLE:
            message = "SVG support disabled. Install cairosvg and pillow to show icons."
            if SVG_ERROR:
                message = f"SVG support disabled: {SVG_ERROR}"
            self.status_var.set(message)

    def sequence_for(self, name: str) -> str:
        strat = self.stratagem_map.get(name)
        if not strat:
            return "Sequence: ?"
        return "Sequence: " + " ".join(strat.sequence)

    def on_stratagem_change(self, index: int) -> None:
        self.equipped[index] = self.slot_vars[index].get()
        self.sequence_labels[index].configure(text=self.sequence_for(self.equipped[index]))
        self.update_icon(index)
        save_user_data(self.equipped, self.keybinds)

    def update_icon(self, index: int) -> None:
        name = self.equipped[index]
        svg_path = ICON_DIR / f"{name}.svg"
        size = 56
        cache_key = (name, size)

        if cache_key in self.icon_cache and self.icon_cache[cache_key]:
            photo = self.icon_cache[cache_key]
            self.icon_labels[index].configure(image=photo, text="")
            self.icon_labels[index].image = photo
            return

        self.icon_labels[index].configure(image="", text="Loading...", fg=MUTED_FG, font=("Segoe UI", 8))

        if not svg_path.exists() or not SVG_AVAILABLE:
            self.icon_labels[index].configure(text="No Icon")
            return

        if cache_key in self.icon_jobs:
            return

        self.icon_jobs.add(cache_key)
        threading.Thread(
            target=self.render_icon_async,
            args=(name, size, svg_path),
            daemon=True,
        ).start()

    def render_icon_async(self, name: str, size: int, svg_path: Path) -> None:
        cache_key = (name, size)
        cache_path: Optional[Path] = None
        error_message: Optional[str] = None
        try:
            cache_path = render_svg_to_png(svg_path, size)
        except Exception as exc:
            error_message = f"Icon render failed: {exc}"
            cache_path = None

        def apply_icon() -> None:
            try:
                if cache_path and cache_path.exists():
                    image = Image.open(cache_path).convert("RGBA")
                    photo = ImageTk.PhotoImage(image)
                    self.icon_cache[cache_key] = photo
                elif error_message and error_message != self.last_icon_error:
                    self.last_icon_error = error_message
                    self.status_var.set(error_message)
            finally:
                self.icon_jobs.discard(cache_key)
                for idx, current_name in enumerate(self.equipped):
                    if current_name == name:
                        self.update_icon(idx)

        self.root.after(0, apply_icon)

    def register_hotkeys(self) -> None:
        if os.name != "nt":
            self.status_var.set("Hotkeys are supported only on Windows.")
            return
        try:
            self.hotkeys = HotkeyManager()
        except RuntimeError as exc:
            self.status_var.set(str(exc))
            return

        for idx, keybind in enumerate(self.keybinds):
            vk = parse_key_code(keybind["key_code"])
            hotkey_id = 1000 + idx
            try:
                self.hotkeys.register(hotkey_id, vk, lambda i=idx: self.activate_stratagem(i))
            except RuntimeError as exc:
                self.status_var.set(str(exc))

        self.root.after(25, self.poll_hotkeys)
        self.root.protocol("WM_DELETE_WINDOW", self.on_close)

    def poll_hotkeys(self) -> None:
        if self.hotkeys:
            self.hotkeys.pump()
        self.root.after(25, self.poll_hotkeys)

    def activate_stratagem(self, index: int) -> None:
        name = self.equipped[index]
        strat = self.stratagem_map.get(name)
        if not strat:
            self.status_var.set(f"Unknown stratagem: {name}")
            return
        self.status_var.set(f"Activated: {name}")
        threading.Thread(target=self.send_sequence, args=(strat.sequence,), daemon=True).start()

    def send_sequence(self, sequence: List[str]) -> None:
        import ctypes

        if os.name != "nt":
            return
        user32 = ctypes.windll.user32
        for entry in sequence:
            vk = KEY_VK.get(entry.upper())
            if not vk:
                continue
            user32.keybd_event(vk, 0, 0, 0)
            time.sleep(0.02)
            user32.keybd_event(vk, 0, 0x0002, 0)
            time.sleep(0.04)

    def on_close(self) -> None:
        if self.hotkeys:
            self.hotkeys.unregister_all()
        self.root.destroy()


def main() -> None:
    root = tk.Tk()
    style = ttk.Style()
    try:
        style.theme_use("clam")
    except Exception:
        pass
    app = StratagemApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
