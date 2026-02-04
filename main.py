from __future__ import annotations

import json
import os
import queue
import sys
import threading
import time
from dataclasses import dataclass
import math
from pathlib import Path
from typing import Callable, Dict, List, Optional, Tuple

import tkinter as tk
from tkinter import messagebox, simpledialog, ttk
# Icons and stratagem info.
# https://github.com/nvigneux/Helldivers-2-Stratagems-icons-svg
# https://helldivers.wiki.gg/wiki/Category:Stratagems

SVG_ERROR = None
try:
    from reportlab.graphics import renderPM  # type: ignore
    from svglib.svglib import svg2rlg  # type: ignore
    from PIL import Image, ImageTk  # type: ignore

    SVG_AVAILABLE = True
except Exception as exc:
    SVG_AVAILABLE = False
    SVG_ERROR = str(exc)


APP_TITLE = "Stratagem Hotkeys"
BASE_DIR = Path(__file__).resolve().parent
RESOURCE_DIR = Path(getattr(sys, "_MEIPASS", BASE_DIR))
USER_DATA_DIR = Path(os.getenv("APPDATA", str(BASE_DIR))) / "pystrat"
USER_DATA_DIR.mkdir(parents=True, exist_ok=True)
DATA_FILE = USER_DATA_DIR / "user_data.json"
STRATAGEMS_FILE = RESOURCE_DIR / "stratagems.json"
ICON_DIR = RESOURCE_DIR / "StratagemIcons"
ICON_CACHE_DIR = USER_DATA_DIR / ".icon_cache"

DARK_BG = "#131316"
CARD_BG = "#1c1c22"
TEXT_FG = "#f1f1f3"
MUTED_FG = "#a2a2ad"
SCROLLBAR_BG = "#2a2a33"
SCROLLBAR_ACTIVE_BG = "#343441"
ICON_RENDER_TAG = "dark_bg"

KEY_VK = {
    "W": 0x57,
    "A": 0x41,
    "S": 0x53,
    "D": 0x44,
}

WM_HOTKEY = 0x0312
WM_QUIT = 0x0012
MOD_NOREPEAT = 0x4000
DEBUG_KEY_CAPTURE = False

LOCAL_KEYSYM_TO_INDEX = {
    "KP_7": 0,
    "KP_8": 1,
    "KP_9": 2,
    "KP_4": 3,
    "KP_5": 4,
    "KP_6": 5,
    "KP_1": 6,
    "KP_2": 7,
    "KP_3": 8,
}


@dataclass
class Stratagem:
    name: str
    sequence: List[str]
    category: str


@dataclass
class UserData:
    equipped_stratagems: List[str]
    keybinds: List[Dict[str, str]]
    key_delay_ms: int
    presets: Dict[str, List[str]]
    active_preset: str
    input_keys: str

    @classmethod
    def load(cls, path: Path) -> "UserData":
        if not path.exists():
            return cls(
                equipped_stratagems=[],
                keybinds=[],
                key_delay_ms=40,
                presets={},
                active_preset="",
                input_keys="wasd",
            )
        raw = json.loads(path.read_text(encoding="utf-8"))
        input_keys = raw.get("input_keys", "wasd")
        if input_keys not in ("wasd", "arrows"):
            input_keys = "wasd"
        return cls(
            equipped_stratagems=raw.get("equipped_stratagems", []),
            keybinds=raw.get("keybinds", []),
            key_delay_ms=int(raw.get("key_delay_ms", 40)),
            presets=raw.get("presets", {}),
            active_preset=raw.get("active_preset", ""),
            input_keys=input_keys,
        )

    def to_payload(self) -> Dict[str, object]:
        return {
            "equipped_stratagems": self.equipped_stratagems,
            "keybinds": self.keybinds,
            "key_delay_ms": self.key_delay_ms,
            "presets": self.presets,
            "active_preset": self.active_preset,
            "input_keys": self.input_keys,
        }

    def save(self, path: Path) -> None:
        path.write_text(json.dumps(self.to_payload(), indent=4), encoding="utf-8")


def load_stratagems() -> List[Stratagem]:
    with STRATAGEMS_FILE.open("r", encoding="utf-8") as handle:
        raw = json.load(handle)
    items: List[Stratagem] = []
    for entry in raw:
        category = entry.get("category", "general")
        items.append(Stratagem(entry["name"], entry["sequence"], category))
    return items


def safe_cache_name(name: str, size: int, tag: str = "") -> str:
    safe = "".join(ch if ch.isalnum() else "_" for ch in name)
    suffix = f"_{tag}" if tag else ""
    return f"{safe}_{size}{suffix}.png"


def render_svg_to_png(svg_path: Path, size: int) -> Optional[Path]:
    if not SVG_AVAILABLE:
        return None
    ICON_CACHE_DIR.mkdir(exist_ok=True)
    cache_path = ICON_CACHE_DIR / safe_cache_name(svg_path.stem, size, ICON_RENDER_TAG)
    if not cache_path.exists() or cache_path.stat().st_mtime < svg_path.stat().st_mtime:
        drawing = svg2rlg(str(svg_path))
        if drawing is None:
            return None
        width = drawing.width or 1
        height = drawing.height or 1
        scale = size / max(width, height)
        drawing.scale(scale, scale)
        drawing.width = width * scale
        drawing.height = height * scale
        png_bytes = renderPM.drawToString(drawing, fmt="PNG", bg=0x0F0F12)
        cache_path.write_bytes(png_bytes)
    return cache_path


def parse_key_code(code: str) -> int:
    return int(code, 16)


class HotkeyManager:
    def __init__(self, notify: Callable[[int], None]) -> None:
        if os.name != "nt":
            raise RuntimeError("Hotkeys require Windows.")
        import ctypes
        from ctypes import wintypes

        self.ctypes = ctypes
        self.wintypes = wintypes
        self.user32 = ctypes.windll.user32
        self.notify = notify
        self.thread: Optional[threading.Thread] = None
        self.thread_id: Optional[int] = None
        self.ready = threading.Event()
        self.errors: List[str] = []
        self.hotkey_map: Dict[int, int] = {}

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

    def start(self, hotkey_map: Dict[int, int]) -> None:
        self.hotkey_map = dict(hotkey_map)
        self.thread = threading.Thread(target=self._message_loop, daemon=True)
        self.thread.start()
        self.ready.wait(timeout=2)

    def stop(self) -> None:
        if self.thread_id is not None:
            self.user32.PostThreadMessageW(self.thread_id, WM_QUIT, 0, 0)
        if self.thread:
            self.thread.join(timeout=1)

    def _message_loop(self) -> None:
        msg = self.MSG()
        self.thread_id = self.ctypes.windll.kernel32.GetCurrentThreadId()
        self.user32.PeekMessageW(self.ctypes.byref(msg), None, 0, 0, 0)

        for hotkey_id, vk in self.hotkey_map.items():
            if not self.user32.RegisterHotKey(None, hotkey_id, MOD_NOREPEAT, vk):
                self.errors.append(f"Failed to register hotkey {hotkey_id}")

        self.ready.set()

        while self.user32.GetMessageW(self.ctypes.byref(msg), None, 0, 0) != 0:
            if msg.message == WM_HOTKEY:
                self.notify(int(msg.wParam))

        for hotkey_id in list(self.hotkey_map.keys()):
            self.user32.UnregisterHotKey(None, hotkey_id)


class StratagemApp:
    def __init__(self, root: tk.Tk) -> None:
        self.root = root
        self.root.title(APP_TITLE)
        self.root.configure(bg=DARK_BG)
        icon_path = RESOURCE_DIR / "app.ico"
        if icon_path.exists():
            self.root.iconbitmap(default=str(icon_path))
        self.root.geometry("820x540")

        self.stratagems = load_stratagems()
        self.stratagem_map = {item.name: item for item in self.stratagems}
        self.stratagem_category = {item.name: item.category for item in self.stratagems}
        self.stratagem_names = [item.name for item in self.stratagems]

        self.user_data = UserData.load(DATA_FILE)
        self.keybinds = self.user_data.keybinds
        self.equipped = self.user_data.equipped_stratagems
        self.key_delay_ms = self.user_data.key_delay_ms
        self.presets = self.user_data.presets
        self.active_preset = self.user_data.active_preset
        self.input_keys = self.user_data.input_keys

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
        self.persist_user_data()

        self.icon_cache: Dict[Tuple[str, int], Optional["ImageTk.PhotoImage"]] = {}
        self.icon_jobs: set[Tuple[str, int]] = set()
        self.sequence_labels: List[tk.Label] = []
        self.icon_labels: List[tk.Label] = []
        self.name_labels: List[tk.Label] = []

        self.status_var = tk.StringVar(value="Ready")
        self.last_icon_error: Optional[str] = None

        self.hotkeys: Optional[HotkeyManager] = None
        self.hotkey_id_to_index: Dict[int, int] = {}
        self.ui_queue: "queue.Queue[Callable[[], None]]" = queue.Queue()
        self.build_ui()
        self.register_hotkeys()
        self.register_local_bindings()
        if DEBUG_KEY_CAPTURE:
            self.register_debug_key_capture()
        self.root.after(100, self.root.focus_set)
        self.root.after(30, self.process_ui_queue)

    def build_ui(self) -> None:
        self.root.grid_rowconfigure(2, weight=1)
        self.root.grid_columnconfigure(0, weight=1)

        header = tk.Label(
            self.root,
            text="Stratagem Hotkeys",
            bg=DARK_BG,
            fg=TEXT_FG,
            font=("Segoe UI", 18, "bold"),
        )
        header.grid(row=0, column=0, sticky="n", pady=(18, 6))

        subtitle = tk.Label(
            self.root,
            text="Press a numpad key to activate the assigned stratagem.",
            bg=DARK_BG,
            fg=MUTED_FG,
            font=("Segoe UI", 10),
        )
        subtitle.grid(row=1, column=0, sticky="n", pady=(0, 12))

        delay_frame = tk.Frame(self.root, bg=DARK_BG)
        delay_frame.grid(row=1, column=0, sticky="w", padx=20, pady=(0, 12))
        delay_label = tk.Label(
            delay_frame,
            text="Key Delay (ms):",
            bg=DARK_BG,
            fg=MUTED_FG,
            font=("Segoe UI", 9),
        )
        delay_label.pack(side="left", padx=(0, 6))
        self.key_delay_var = tk.IntVar(value=self.key_delay_ms)
        delay_spin = tk.Spinbox(
            delay_frame,
            from_=10,
            to=300,
            increment=5,
            textvariable=self.key_delay_var,
            width=5,
            command=self.on_key_delay_change,
            justify="center",
        )
        delay_spin.bind("<FocusOut>", lambda _e: self.on_key_delay_change())
        delay_spin.pack(side="left")

        key_mode_frame = tk.Frame(self.root, bg=DARK_BG)
        key_mode_frame.grid(row=1, column=0, sticky="w", padx=180, pady=(0, 12))
        key_mode_label = tk.Label(
            key_mode_frame,
            text="Input Keys:",
            bg=DARK_BG,
            fg=MUTED_FG,
            font=("Segoe UI", 9),
        )
        key_mode_label.pack(side="left", padx=(0, 6))
        self.input_keys_var = tk.StringVar(value="WASD" if self.input_keys == "wasd" else "Arrows")
        key_mode_combo = ttk.Combobox(
            key_mode_frame,
            textvariable=self.input_keys_var,
            values=["WASD", "Arrows"],
            state="readonly",
            width=8,
        )
        key_mode_combo.pack(side="left")
        key_mode_combo.bind("<<ComboboxSelected>>", self.on_input_keys_change)

        preset_frame = tk.Frame(self.root, bg=DARK_BG)
        preset_frame.grid(row=0, column=0, sticky="e", padx=20, pady=(18, 6))
        preset_label = tk.Label(
            preset_frame,
            text="Preset:",
            bg=DARK_BG,
            fg=MUTED_FG,
            font=("Segoe UI", 9),
        )
        preset_label.pack(side="left", padx=(0, 6))
        self.preset_var = tk.StringVar(value=self.active_preset)
        self.preset_combo = ttk.Combobox(
            preset_frame,
            textvariable=self.preset_var,
            values=self.get_preset_names(),
            state="readonly",
            width=20,
        )
        self.preset_combo.pack(side="left", padx=(0, 6))
        self.preset_combo.bind("<<ComboboxSelected>>", self.on_preset_select)
        preset_save = ttk.Button(preset_frame, text="Save", command=self.save_preset_current)
        preset_save.pack(side="left")
        preset_menu_button = tk.Menubutton(
            preset_frame,
            text="▾",
            bg=CARD_BG,
            fg=TEXT_FG,
            activebackground=CARD_BG,
            activeforeground=TEXT_FG,
            relief="ridge",
            font=("Segoe UI Black", 12, "bold"),
            width=2,
        )
        preset_menu = tk.Menu(preset_menu_button, tearoff=0)
        preset_menu.add_command(label="Save As...", command=self.save_preset_prompt)
        preset_menu.add_separator()
        preset_menu.add_command(label="Delete", command=self.delete_preset)
        preset_menu_button.configure(menu=preset_menu)
        preset_menu_button.pack(side="left", padx=(2, 0))

        grid_frame = tk.Frame(self.root, bg=DARK_BG)
        grid_frame.grid(row=2, column=0, sticky="nsew", padx=20, pady=10)

        columns = 3
        rows = 3
        for col in range(columns):
            grid_frame.grid_columnconfigure(col, weight=1, uniform="strat_cols")
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
            icon_label.bind("<Button-1>", lambda _e, i=index: self.open_icon_picker(i))

            name_label = tk.Label(
                card,
                text=self.equipped[index],
                bg=CARD_BG,
                fg=TEXT_FG,
                font=("Segoe UI", 11, "bold"),
                anchor="w",
                cursor="hand2",
            )
            name_label.grid(row=1, column=1, sticky="ew")
            name_label.bind("<Button-1>", lambda _e, i=index: self.open_icon_picker(i))

            seq_label = tk.Label(
                card,
                text=self.sequence_for(self.equipped[index]),
                bg=CARD_BG,
                fg=MUTED_FG,
                font=("Segoe UI Black", 14, "bold"),
            )
            seq_label.grid(row=2, column=1, sticky="w")

            self.sequence_labels.append(seq_label)
            self.icon_labels.append(icon_label)
            self.name_labels.append(name_label)

            self.update_icon(index)

        status_frame = tk.Frame(self.root, bg=CARD_BG, height=28)
        status_frame.grid(row=3, column=0, sticky="ew")
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
            message = "SVG support disabled. Install svglib, reportlab, and pillow to show icons."
            if SVG_ERROR:
                message = f"SVG support disabled: {SVG_ERROR}"
            self.status_var.set(message)

    def persist_user_data(self) -> None:
        self.user_data.equipped_stratagems = list(self.equipped)
        self.user_data.keybinds = list(self.keybinds)
        self.user_data.key_delay_ms = self.key_delay_ms
        self.user_data.presets = dict(self.presets)
        self.user_data.active_preset = self.active_preset
        self.user_data.input_keys = self.input_keys
        self.user_data.save(DATA_FILE)

    def sequence_for(self, name: str) -> str:
        strat = self.stratagem_map.get(name)
        if not strat:
            return "?"
        arrows = {"W": "⬆", "A": "⬅", "S": "⬇", "D": "➡"}
        display = [arrows.get(step.upper(), step) for step in strat.sequence]
        return " ".join(display)

    def set_stratagem(self, index: int, name: str) -> None:
        self.equipped[index] = name
        self.sequence_labels[index].configure(text=self.sequence_for(name))
        self.name_labels[index].configure(text=name)
        self.update_icon(index)
        self.persist_user_data()

    def open_icon_picker(self, index: int) -> None:
        picker = tk.Toplevel(self.root)
        picker.title("Select Stratagem")
        picker.configure(bg=DARK_BG)
        picker.transient(self.root)
        picker.grab_set()
        picker.update_idletasks()
        width = 720
        height = 520
        root_x = self.root.winfo_rootx()
        root_y = self.root.winfo_rooty()
        root_w = self.root.winfo_width()
        root_h = self.root.winfo_height()
        pos_x = max(root_x + (root_w - width) // 2, 0)
        pos_y = max(root_y + (root_h - height) // 2, 0)
        picker.geometry(f"{width}x{height}+{pos_x}+{pos_y}")

        header = tk.Label(
            picker,
            text="Select Stratagem",
            bg=DARK_BG,
            fg=TEXT_FG,
            font=("Segoe UI", 14, "bold"),
        )
        header.pack(pady=(16, 8))

        search_frame = tk.Frame(picker, bg=DARK_BG)
        search_frame.pack(fill="x", padx=16, pady=(0, 8))
        search_label = tk.Label(
            search_frame,
            text="Search:",
            bg=DARK_BG,
            fg=MUTED_FG,
            font=("Segoe UI", 9),
        )
        search_label.pack(side="left", padx=(0, 6))
        search_var = tk.StringVar()
        search_entry = tk.Entry(search_frame, textvariable=search_var, width=28)
        search_entry.pack(side="left")
        search_entry.focus_set()

        container = tk.Frame(picker, bg=DARK_BG)
        container.pack(fill="both", expand=True, padx=16, pady=(0, 16))

        canvas = tk.Canvas(container, bg=DARK_BG, highlightthickness=0)
        scrollbar = ttk.Scrollbar(
            container,
            orient="vertical",
            command=canvas.yview,
            style="Strat.Vertical.TScrollbar",
        )
        scroll_frame = tk.Frame(canvas, bg=DARK_BG)
        scroll_frame.bind(
            "<Configure>",
            lambda _e: canvas.configure(scrollregion=canvas.bbox("all")),
        )
        canvas.create_window((0, 0), window=scroll_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
        canvas.bind_all("<MouseWheel>", lambda e: canvas.yview_scroll(int(-1 * (e.delta / 120)), "units"))
        canvas.bind_all("<Button-4>", lambda _e: canvas.yview_scroll(-3, "units"))
        canvas.bind_all("<Button-5>", lambda _e: canvas.yview_scroll(3, "units"))

        columns = 4
        size = 52
        category_order = ["Offensive", "Supply", "Defensive", "General"]
        category_labels = {
            "Offensive": "Offensive",
            "Supply": "Supply",
            "Defensive": "Defensive",
            "General": "General",
        }

        def rebuild_grid(*_args: object) -> None:
            for child in scroll_frame.winfo_children():
                child.destroy()

            filter_text = search_var.get().strip().lower()
            names = self.stratagem_names
            if filter_text:
                names = [name for name in names if filter_text in name.lower()]

            row_cursor = 0
            for category in category_order:
                cat_names = [name for name in names if self.stratagem_category.get(name) == category]
                if not cat_names:
                    continue

                header = tk.Label(
                    scroll_frame,
                    text=category_labels.get(category, category.title()),
                    bg=DARK_BG,
                    fg=MUTED_FG,
                    font=("Segoe UI", 10, "bold"),
                    anchor="w",
                )
                header.grid(row=row_cursor, column=0, columnspan=columns, sticky="w", padx=6, pady=(10, 4))
                row_cursor += 1

                for idx, name in enumerate(cat_names):
                    row = row_cursor + idx // columns
                    col = idx % columns
                    cell = tk.Frame(scroll_frame, bg=CARD_BG, padx=6, pady=6)
                    cell.grid(row=row, column=col, padx=6, pady=6, sticky="nsew")

                    icon = tk.Label(cell, bg="#0f0f12", width=size, height=size)
                    icon.pack()
                    photo = self.get_icon_photo(name, size)
                    if photo:
                        icon.configure(image=photo)
                        icon.image = photo
                    else:
                        icon.configure(text="No Icon", fg=MUTED_FG, font=("Segoe UI", 7))

                    label = tk.Label(
                        cell,
                        text=name,
                        bg=CARD_BG,
                        fg=TEXT_FG,
                        font=("Segoe UI", 8),
                        wraplength=110,
                        justify="center",
                    )
                    label.pack(pady=(4, 0))

                    for widget in (cell, icon, label):
                        widget.bind(
                            "<Button-1>",
                            lambda _e, n=name: self._select_stratagem_from_picker(picker, index, n),
                        )

                row_cursor += int(math.ceil(len(cat_names) / columns))

        search_var.trace_add("write", lambda *_a: rebuild_grid())
        rebuild_grid()

    def _select_stratagem_from_picker(self, picker: tk.Toplevel, index: int, name: str) -> None:
        self.set_stratagem(index, name)
        picker.destroy()

    def get_icon_photo(self, name: str, size: int) -> Optional["ImageTk.PhotoImage"]:
        cache_key = (name, size)
        if cache_key in self.icon_cache and self.icon_cache[cache_key]:
            return self.icon_cache[cache_key]

        svg_path = ICON_DIR / f"{name}.svg"
        if not svg_path.exists() or not SVG_AVAILABLE:
            return None
        try:
            cache_path = render_svg_to_png(svg_path, size)
            if not cache_path or not cache_path.exists():
                return None
            image = Image.open(cache_path).convert("RGBA")
            photo = ImageTk.PhotoImage(image)
            self.icon_cache[cache_key] = photo
            return photo
        except Exception:
            return None

    def on_key_delay_change(self) -> None:
        try:
            value = int(self.key_delay_var.get())
        except Exception:
            return
        value = max(10, min(300, value))
        self.key_delay_ms = value
        self.key_delay_var.set(value)
        self.persist_user_data()

    def on_input_keys_change(self, _event: tk.Event) -> None:
        label = self.input_keys_var.get().strip().lower()
        self.input_keys = "arrows" if "arrow" in label else "wasd"
        self.persist_user_data()

    def get_preset_names(self) -> List[str]:
        return sorted(self.presets.keys())

    def refresh_preset_combo(self) -> None:
        self.preset_combo.configure(values=self.get_preset_names())
        if self.active_preset:
            self.preset_var.set(self.active_preset)

    def save_preset_prompt(self) -> None:
        name = simpledialog.askstring("Save Preset", "Preset name:", parent=self.root)
        if not name:
            return
        self.presets[name] = list(self.equipped)
        self.active_preset = name
        self.refresh_preset_combo()
        self.persist_user_data()

    def save_preset_current(self) -> None:
        if not self.active_preset:
            self.save_preset_prompt()
            return
        self.presets[self.active_preset] = list(self.equipped)
        self.persist_user_data()

    def delete_preset(self) -> None:
        name = self.preset_var.get()
        if not name:
            return
        if not messagebox.askyesno("Delete Preset", f"Delete preset '{name}'?", parent=self.root):
            return
        self.presets.pop(name, None)
        if self.active_preset == name:
            self.active_preset = ""
            self.preset_var.set("")
        self.refresh_preset_combo()
        self.persist_user_data()

    def on_preset_select(self, _event: tk.Event) -> None:
        name = self.preset_var.get()
        if not name:
            return
        loadout = self.presets.get(name)
        if not loadout:
            return
        self.active_preset = name
        for idx, strat_name in enumerate(loadout[: len(self.equipped)]):
            self.set_stratagem(idx, strat_name)
        self.persist_user_data()

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

        self.run_in_ui(apply_icon)

    def register_hotkeys(self) -> None:
        if os.name != "nt":
            self.status_var.set("Hotkeys are supported only on Windows.")
            return
        try:
            self.hotkeys = HotkeyManager(self.on_hotkey_fired)
        except RuntimeError as exc:
            self.status_var.set(str(exc))
            return

        hotkey_map: Dict[int, int] = {}
        for idx, keybind in enumerate(self.keybinds):
            vk = parse_key_code(keybind["key_code"])
            hotkey_id = 1000 + idx
            self.hotkey_id_to_index[hotkey_id] = idx
            hotkey_map[hotkey_id] = vk

        self.hotkeys.start(hotkey_map)
        if self.hotkeys.errors:
            self.status_var.set(self.hotkeys.errors[0])

        self.root.protocol("WM_DELETE_WINDOW", self.on_close)

    def register_local_bindings(self) -> None:
        for keysym, index in LOCAL_KEYSYM_TO_INDEX.items():
            self.root.bind(f"<KeyPress-{keysym}>", lambda _e, i=index: self.activate_stratagem(i))

    def register_debug_key_capture(self) -> None:
        self.root.bind_all("<KeyPress>", self.on_any_key)

    def on_any_key(self, event: tk.Event) -> None:
        keysym = getattr(event, "keysym", "")
        keycode = getattr(event, "keycode", "")
        char = getattr(event, "char", "")
        self.status_var.set(f"Key: keysym={keysym} keycode={keycode} char={char}")

    def on_hotkey_fired(self, hotkey_id: int) -> None:
        index = self.hotkey_id_to_index.get(hotkey_id)
        if index is None:
            return
        self.run_in_ui(lambda: self.activate_stratagem(index))

    def run_in_ui(self, func: Callable[[], None]) -> None:
        self.ui_queue.put(func)

    def process_ui_queue(self) -> None:
        while True:
            try:
                func = self.ui_queue.get_nowait()
            except queue.Empty:
                break
            try:
                if self.root.winfo_exists():
                    func()
            except Exception as exc:
                self.status_var.set(f"UI update error: {exc}")
        if self.root.winfo_exists():
            self.root.after(30, self.process_ui_queue)

    def activate_stratagem(self, index: int) -> None:
        name = self.equipped[index]
        strat = self.stratagem_map.get(name)
        if not strat:
            self.status_var.set(f"Unknown stratagem: {name}")
            return
        sequence_text = " ".join(strat.sequence)
        self.status_var.set(f"Activated: {name} ({sequence_text})")
        threading.Thread(target=self.send_sequence, args=(strat.sequence,), daemon=True).start()

    def send_sequence(self, sequence: List[str]) -> None:
        import ctypes

        if os.name != "nt":
            return

        user32 = ctypes.windll.user32
        KEYEVENTF_SCANCODE = 0x0008
        KEYEVENTF_KEYUP = 0x0002
        VK_CONTROL = 0x11
        VK_UP = 0x26
        VK_LEFT = 0x25
        VK_DOWN = 0x28
        VK_RIGHT = 0x27

        class KEYBDINPUT(ctypes.Structure):
            _fields_ = [
                ("wVk", ctypes.c_ushort),
                ("wScan", ctypes.c_ushort),
                ("dwFlags", ctypes.c_ulong),
                ("time", ctypes.c_ulong),
                ("dwExtraInfo", ctypes.c_void_p),
            ]

        class MOUSEINPUT(ctypes.Structure):
            _fields_ = [
                ("dx", ctypes.c_long),
                ("dy", ctypes.c_long),
                ("mouseData", ctypes.c_ulong),
                ("dwFlags", ctypes.c_ulong),
                ("time", ctypes.c_ulong),
                ("dwExtraInfo", ctypes.c_void_p),
            ]

        class HARDWAREINPUT(ctypes.Structure):
            _fields_ = [
                ("uMsg", ctypes.c_ulong),
                ("wParamL", ctypes.c_ushort),
                ("wParamH", ctypes.c_ushort),
            ]

        class INPUTUNION(ctypes.Union):
            _fields_ = [
                ("ki", KEYBDINPUT),
                ("mi", MOUSEINPUT),
                ("hi", HARDWAREINPUT),
            ]

        class INPUT(ctypes.Structure):
            _fields_ = [("type", ctypes.c_ulong), ("union", INPUTUNION)]

        user32.SendInput.argtypes = [ctypes.c_uint, ctypes.POINTER(INPUT), ctypes.c_int]
        user32.SendInput.restype = ctypes.c_uint

        def send_key(vk: int, flags: int) -> None:
            inp = INPUT(1, INPUTUNION(ki=KEYBDINPUT(vk, 0, flags, 0, None)))
            user32.SendInput(1, ctypes.byref(inp), ctypes.sizeof(INPUT))

        def send_ctrl(flags: int) -> None:
            scan = user32.MapVirtualKeyW(VK_CONTROL, 0)
            inp = INPUT(1, INPUTUNION(ki=KEYBDINPUT(0, scan, flags | KEYEVENTF_SCANCODE, 0, None)))
            user32.SendInput(1, ctypes.byref(inp), ctypes.sizeof(INPUT))

        send_ctrl(0)
        time.sleep(0.02)
        arrow_vk = {"W": VK_UP, "A": VK_LEFT, "S": VK_DOWN, "D": VK_RIGHT}

        for entry in sequence:
            key_name = entry.upper()
            if self.input_keys == "arrows":
                vk = arrow_vk.get(key_name)
            else:
                vk = KEY_VK.get(key_name)
            if not vk:
                continue
            send_key(vk, 0)
            time.sleep(0.02)
            send_key(vk, KEYEVENTF_KEYUP)
            time.sleep(self.key_delay_ms / 1000.0)
        send_ctrl(KEYEVENTF_KEYUP)

    def on_close(self) -> None:
        if self.hotkeys:
            self.hotkeys.stop()
        self.root.destroy()


def main() -> None:
    root = tk.Tk()
    style = ttk.Style()
    try:
        style.theme_use("clam")
    except tk.TclError:
        pass
    style.configure(
        "Strat.Vertical.TScrollbar",
        troughcolor=DARK_BG,
        background=SCROLLBAR_BG,
        bordercolor=DARK_BG,
        lightcolor=SCROLLBAR_BG,
        darkcolor=SCROLLBAR_BG,
        arrowcolor=MUTED_FG,
        gripcount=0,
        width=10,
    )
    style.map(
        "Strat.Vertical.TScrollbar",
        background=[("active", SCROLLBAR_ACTIVE_BG), ("pressed", SCROLLBAR_ACTIVE_BG)],
    )
    app = StratagemApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
