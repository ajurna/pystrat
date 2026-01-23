from __future__ import annotations

import json
import os
import queue
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
WM_QUIT = 0x0012
MOD_NOREPEAT = 0x4000
DEBUG_KEY_CAPTURE = False

INPUT_MODE_OPTIONS = {
    "scancode": "Scancode",
    "vk": "Virtual Key",
    "unicode": "Unicode",
}

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


def load_stratagems() -> List[Stratagem]:
    with STRATAGEMS_FILE.open("r", encoding="utf-8") as handle:
        raw = json.load(handle)
    return [Stratagem(entry["name"], entry["sequence"]) for entry in raw]


def load_user_data() -> Dict[str, List]:
    if not DATA_FILE.exists():
        return {"equipped_stratagems": [], "keybinds": [], "key_delay_ms": 40}
    with DATA_FILE.open("r", encoding="utf-8") as handle:
        return json.load(handle)


def save_user_data(
    equipped: List[str],
    keybinds: List[Dict[str, str]],
    input_mode: Optional[str] = None,
    key_delay_ms: Optional[int] = None,
) -> None:
    payload = {"equipped_stratagems": equipped, "keybinds": keybinds}
    if input_mode:
        payload["input_mode"] = input_mode
    if key_delay_ms is not None:
        payload["key_delay_ms"] = key_delay_ms
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
        self.root.geometry("820x540")

        self.stratagems = load_stratagems()
        self.stratagem_map = {item.name: item for item in self.stratagems}
        self.stratagem_names = [item.name for item in self.stratagems]

        self.user_data = load_user_data()
        self.keybinds = self.user_data.get("keybinds", [])
        self.equipped = self.user_data.get("equipped_stratagems", [])
        self.input_mode = self.user_data.get("input_mode", "scancode")
        self.key_delay_ms = int(self.user_data.get("key_delay_ms", 40))
        if self.input_mode not in INPUT_MODE_OPTIONS:
            self.input_mode = "scancode"

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
        save_user_data(self.equipped, self.keybinds, self.input_mode, self.key_delay_ms)

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

        mode_frame = tk.Frame(self.root, bg=DARK_BG)
        mode_frame.grid(row=1, column=0, sticky="e", padx=20, pady=(0, 12))
        mode_label = tk.Label(
            mode_frame,
            text="Input Mode:",
            bg=DARK_BG,
            fg=MUTED_FG,
            font=("Segoe UI", 9),
        )
        mode_label.pack(side="left", padx=(0, 6))

        self.input_mode_var = tk.StringVar(value=INPUT_MODE_OPTIONS[self.input_mode])
        mode_combo = ttk.Combobox(
            mode_frame,
            textvariable=self.input_mode_var,
            values=list(INPUT_MODE_OPTIONS.values()),
            state="readonly",
            width=12,
        )
        mode_combo.pack(side="left")
        mode_combo.bind("<<ComboboxSelected>>", self.on_input_mode_change)

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

        grid_frame = tk.Frame(self.root, bg=DARK_BG)
        grid_frame.grid(row=2, column=0, sticky="nsew", padx=20, pady=10)

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
                font=("Consolas", 11),
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
            message = "SVG support disabled. Install cairosvg and pillow to show icons."
            if SVG_ERROR:
                message = f"SVG support disabled: {SVG_ERROR}"
            self.status_var.set(message)

    def sequence_for(self, name: str) -> str:
        strat = self.stratagem_map.get(name)
        if not strat:
            return "?"
        return " ".join(strat.sequence)

    def set_stratagem(self, index: int, name: str) -> None:
        self.equipped[index] = name
        self.sequence_labels[index].configure(text=self.sequence_for(name))
        self.name_labels[index].configure(text=name)
        self.update_icon(index)
        save_user_data(self.equipped, self.keybinds, self.input_mode, self.key_delay_ms)

    def open_icon_picker(self, index: int) -> None:
        picker = tk.Toplevel(self.root)
        picker.title("Select Stratagem")
        picker.configure(bg=DARK_BG)
        picker.geometry("720x520")
        picker.transient(self.root)
        picker.grab_set()

        header = tk.Label(
            picker,
            text="Select Stratagem",
            bg=DARK_BG,
            fg=TEXT_FG,
            font=("Segoe UI", 14, "bold"),
        )
        header.pack(pady=(16, 8))

        container = tk.Frame(picker, bg=DARK_BG)
        container.pack(fill="both", expand=True, padx=16, pady=(0, 16))

        canvas = tk.Canvas(container, bg=DARK_BG, highlightthickness=0)
        scrollbar = ttk.Scrollbar(container, orient="vertical", command=canvas.yview)
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

        columns = 6
        size = 52
        for idx, name in enumerate(self.stratagem_names):
            row = idx // columns
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

    def on_input_mode_change(self, _event: tk.Event) -> None:
        label = self.input_mode_var.get()
        reverse_map = {value: key for key, value in INPUT_MODE_OPTIONS.items()}
        self.input_mode = reverse_map.get(label, "scancode")
        save_user_data(self.equipped, self.keybinds, self.input_mode, self.key_delay_ms)

    def on_key_delay_change(self) -> None:
        try:
            value = int(self.key_delay_var.get())
        except Exception:
            return
        value = max(10, min(300, value))
        self.key_delay_ms = value
        self.key_delay_var.set(value)
        save_user_data(self.equipped, self.keybinds, self.input_mode, self.key_delay_ms)

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
        KEYEVENTF_UNICODE = 0x0004
        VK_CONTROL = 0x11

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
            if self.input_mode == "unicode":
                inp = INPUT(1, INPUTUNION(ki=KEYBDINPUT(0, vk, flags | KEYEVENTF_UNICODE, 0, None)))
            elif self.input_mode == "vk":
                inp = INPUT(1, INPUTUNION(ki=KEYBDINPUT(vk, 0, flags, 0, None)))
            else:
                scan = user32.MapVirtualKeyW(vk, 0)
                inp = INPUT(1, INPUTUNION(ki=KEYBDINPUT(0, scan, flags | KEYEVENTF_SCANCODE, 0, None)))
            user32.SendInput(1, ctypes.byref(inp), ctypes.sizeof(INPUT))

        def send_ctrl(flags: int) -> None:
            scan = user32.MapVirtualKeyW(VK_CONTROL, 0)
            inp = INPUT(1, INPUTUNION(ki=KEYBDINPUT(0, scan, flags | KEYEVENTF_SCANCODE, 0, None)))
            user32.SendInput(1, ctypes.byref(inp), ctypes.sizeof(INPUT))

        send_ctrl(0)
        time.sleep(0.02)
        for entry in sequence:
            key_name = entry.upper()
            vk = KEY_VK.get(key_name)
            if not vk:
                continue
            if self.input_mode == "unicode":
                code = ord(key_name.lower())
                send_key(code, 0)
                time.sleep(0.02)
                send_key(code, KEYEVENTF_KEYUP)
            else:
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
    except Exception:
        pass
    app = StratagemApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
