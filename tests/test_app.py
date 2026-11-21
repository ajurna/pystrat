import base64
import ctypes
import json
import os
import runpy
import sys
import tempfile
import time
import tkinter as tk
import unittest
from pathlib import Path
from tkinter import ttk
from types import SimpleNamespace
from unittest import mock

import main

# Smallest valid PNG: Tk PhotoImage needs real image data.
TINY_PNG = base64.b64decode(
    "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNk+M9QDwADhgGAWjR9awAAAABJRU5ErkJggg=="
)

FIXTURE_STRATAGEMS = [
    {"name": "Alpha Strike", "sequence": ["W", "A", "S", "D"], "category": "Offensive", "type": "orbital", "cooldown_seconds": 120},
    {"name": "Arc Thrower", "sequence": ["A"], "category": "Offensive", "type": "support_weapon", "cooldown_seconds": 60},
    {"name": "Belt Fed", "sequence": ["X", "Q"], "category": "Supply", "type": "resupply"},
    {"name": "Bullet Camera", "sequence": ["Q"], "category": "Supply", "type": "backpack", "cooldown_seconds": 20},
    {"name": "Eagle Striker", "sequence": ["S"], "category": "Offensive", "type": "eagle", "cooldown_seconds": 30},
    {"name": "Shield", "sequence": ["A"], "category": "Defensive", "type": "support_weapon", "cooldown_seconds": 30},
    {"name": "Shield Generator", "sequence": ["D"], "category": "Defensive", "type": "emplacement", "cooldown_seconds": 90},
    {"name": "Telepylo", "sequence": ["W"], "category": "General", "type": "mission", "cooldown_seconds": 15},
    {"name": "Uncategorised", "sequence": ["E"], "category": "general", "type": "mission"},
]

FIXTURE_NAMES = [entry["name"] for entry in FIXTURE_STRATAGEMS]

DESIRED_LETTERS = [
    "NumPad7",
    "NumPad8",
    "NumPad9",
    "NumPad4",
    "NumPad5",
    "NumPad6",
    "NumPad1",
    "NumPad2",
    "NumPad3",
]


def fake_windll(user32):
    """ctypes.windll stand-in: no real hotkeys, no real keystrokes."""
    return SimpleNamespace(
        user32=user32,
        kernel32=SimpleNamespace(GetCurrentThreadId=lambda: 4242),
    )


def descendants(widget):
    """Every Tk widget under `widget` (ttk widgets included)."""
    for child in widget.winfo_children():
        yield child
        yield from descendants(child)


def exists(widget) -> bool:
    try:
        return bool(widget.winfo_exists())
    except tk.TclError:
        return False


def fake_hotkey_manager(errors=None, raise_runtime=False):
    def factory(notify):
        if raise_runtime:
            raise RuntimeError("Hotkeys require Windows.")
        return SimpleNamespace(
            notify=notify,
            errors=list(errors or []),
            hotkey_map={},
            start=lambda hotkey_map: None,
            stop=lambda: None,
        )

    return factory


class FakeUser32:
    """Records SendInput calls instead of injecting keystrokes."""

    def __init__(self):
        self.sent = []
        # Flags whose SendInput should report failure (0) or raise.
        self.failing: set[int] = set()
        self.raising: set[int] = set()
        # Plain functions (not methods): send_sequence assigns .argtypes/.restype on them.
        self.SendInput = lambda count, inp, size: self.record(count, inp)
        self.MapVirtualKeyW = lambda vk, flags: vk
        self.RegisterHotKey = lambda *a: True
        self.UnregisterHotKey = lambda *a: True
        self.PeekMessageW = lambda *a: 0
        self.PostThreadMessageW = lambda *a: 0
        self.GetMessageW = lambda *a: 0

    def record(self, count, inp):
        flags = int(inp.union.ki.dwFlags)
        self.sent.append((int(inp.union.ki.wVk), flags))
        if flags in self.raising:
            raise RuntimeError("SendInput rejected the key")
        return 0 if flags in self.failing else count


class SyncThread:
    """Runs the target inline so activate_stratagem finishes inside the test."""

    def __init__(self, target=None, args=(), **_kw):
        self.target = target
        self.args = args

    def start(self):
        self.target(*self.args)


class StratagemAppTest(unittest.TestCase):
    def setUp(self):
        tmp = tempfile.TemporaryDirectory()
        self.addCleanup(tmp.cleanup)
        self.tmp = Path(tmp.name)
        self.data_file = self.tmp / "user_data.json"
        self.icon_dir = self.tmp / "icons"
        self.icon_dir.mkdir()
        self.fixture = self.tmp / "stratagems.json"
        self.fixture.write_text(json.dumps(FIXTURE_STRATAGEMS), encoding="utf-8")

        self.user32 = FakeUser32()
        for patcher in (
            mock.patch.object(main, "DATA_FILE", self.data_file),
            mock.patch.object(main, "USER_DATA_DIR", self.tmp),
            mock.patch.object(main, "STRATAGEMS_FILE", self.fixture),
            mock.patch.object(main, "ICON_DIR", self.icon_dir),
            mock.patch.object(main, "render_svg_to_png_bytes", lambda _p, _s: TINY_PNG),
            mock.patch.object(main, "HotkeyManager", fake_hotkey_manager()),
            # Keep the app's self-rescheduling after() calls from firing on a destroyed
            # root (they only print "invalid command name" noise); the closures they
            # schedule are invoked directly by the tests below.
            mock.patch.object(tk.Misc, "after", lambda *_a, **_k: None),
            mock.patch.object(os, "name", "nt"),
            mock.patch.object(ctypes, "windll", fake_windll(self.user32)),
            mock.patch.object(ctypes, "byref", lambda obj: obj),
            mock.patch.object(time, "sleep", lambda _s: None),
            mock.patch.object(main, "threading", SimpleNamespace(Thread=SyncThread)),
        ):
            patcher.start()
            self.addCleanup(patcher.stop)

        self.root = tk.Tk()
        self.root.withdraw()
        self.addCleanup(lambda: exists(self.root) and self.root.destroy())

    def build(self, payload=None):
        if payload is not None:
            self.data_file.write_text(json.dumps(payload), encoding="utf-8")
        return main.StratagemApp(self.root)

    def saved(self):
        return json.loads(self.data_file.read_text(encoding="utf-8"))

    def test_default_user_data_builds_full_ui(self):
        app = self.build()

        self.assertEqual([entry["letter"] for entry in app.keybinds], DESIRED_LETTERS)
        self.assertEqual(len(app.equipped), len(app.keybinds))
        self.assertEqual(app.equipped[0], "Alpha Strike")
        self.assertEqual(app.status_var.get(), "Ready")
        with main.PYPROJECT.open("rb") as handle:
            self.assertEqual(app.version, main.toml_load(handle)["project"]["version"])
        self.assertEqual(len(app.sequence_labels), 9)
        self.assertEqual(len(app.icon_canvases), 9)
        self.assertEqual(app.sequence_labels[0].cget("text"), "⬆ ⬅ ⬇ ➡")
        self.assertEqual(app.cooldown_labels[0].cget("text"), "Ready")
        self.assertEqual(self.saved()["equipped_stratagems"], app.equipped)
        self.assertEqual(self.saved()["key_delay_ms"], 40)

    def test_existing_keybinds_are_merged_and_sorted(self):
        payload = {
            "keybinds": [{"key_code": "0x61", "letter": "NumPad1"}],
            "equipped_stratagems": FIXTURE_NAMES + ["", "Ghost Name"],
            "key_delay_ms": 55,
            "presets": {"loadout": FIXTURE_NAMES[:3]},
            "active_preset": "loadout",
            "input_keys": "arrows",
            "ship_modules": ["hand_carts", "not-a-module"],
        }
        app = self.build(payload)

        self.assertEqual([entry["letter"] for entry in app.keybinds], DESIRED_LETTERS)
        # Names without an icon and a sequence are dropped, not added to the picker.
        self.assertNotIn("Ghost Name", app.stratagem_names)
        self.assertEqual(app.equipped, FIXTURE_NAMES)
        self.assertEqual(app.key_delay_ms, 55)
        self.assertEqual(app.input_keys, "arrows")
        self.assertEqual(app.active_ship_modules, {"hand_carts"})
        self.assertEqual(app.preset_var.get(), "loadout")
        self.assertEqual(
            app.status_var.get(),
            "Fixed settings: dropped unknown ship modules, dropped unknown equipped names",
        )
        # Fixes are only written back on the next change.
        app.persist_user_data()
        self.assertEqual(self.saved()["ship_modules"], ["hand_carts"])

    def test_corrupt_settings_are_not_overwritten(self):
        self.data_file.write_text("{ not json", encoding="utf-8")
        app = self.build()

        self.assertEqual(
            app.status_var.get(),
            "user_data.json is corrupt - using defaults, nothing is saved.",
        )
        self.assertEqual(self.data_file.read_text(encoding="utf-8"), "{ not json")
        self.assertEqual(app.equipped, FIXTURE_NAMES)

    def test_out_of_range_delay_is_clamped(self):
        app = self.build({"key_delay_ms": 5000})
        self.assertEqual(app.key_delay_ms, 300)
        self.assertEqual(app.status_var.get(), "Fixed settings: key delay 5000 -> 300")

    def test_unusable_keybind_entries_are_dropped(self):
        app = self.build(
            {"keybinds": [{"key_code": "zz", "letter": "NumPad1"}, {"letter": "NumPad2"}]}
        )
        self.assertEqual([entry["letter"] for entry in app.keybinds], DESIRED_LETTERS)
        self.assertEqual(
            app.status_var.get(),
            "Fixed settings: dropped unusable keybind entries",
        )

    def test_unknown_preset_slots_are_dropped(self):
        app = self.build({"presets": {"A": ["Shield", "Ghost"]}, "active_preset": "A"})
        self.assertEqual(app.presets, {"A": ["Shield"]})
        self.assertEqual(
            app.status_var.get(),
            "Fixed settings: dropped unknown preset slots",
        )

    def test_empty_preset_and_stale_active_preset_are_dropped(self):
        app = self.build({"presets": {"A": ["Ghost"]}, "active_preset": "A"})
        self.assertEqual(app.presets, {})
        self.assertEqual(app.active_preset, "")
        self.assertEqual(
            app.status_var.get(),
            "Fixed settings: dropped unknown preset slots, active preset 'A' no longer exists",
        )

    def test_reset_corrupt_settings_keeps_the_bad_file(self):
        self.data_file.write_text("{ not json", encoding="utf-8")
        app = self.build()

        app.reset_corrupt_settings()

        backup = Path(str(self.data_file) + ".bak")
        self.assertEqual(backup.read_text(encoding="utf-8"), "{ not json")
        self.assertFalse(app.user_data.corrupt)
        self.assertEqual(self.saved()["equipped_stratagems"], FIXTURE_NAMES)
        self.addCleanup(lambda: backup.unlink(missing_ok=True))

    def test_reset_corrupt_settings_reports_a_failed_backup(self):
        self.data_file.write_text("{ not json", encoding="utf-8")
        app = self.build()
        backup = Path(str(self.data_file) + ".bak")

        with mock.patch.object(Path, "rename", mock.Mock(side_effect=OSError("permission denied"))):
            app.reset_corrupt_settings()

        self.assertEqual(app.status_var.get(), f"Could not move {backup.name}: permission denied")
        self.assertTrue(app.user_data.corrupt)
        self.addCleanup(lambda: backup.unlink(missing_ok=True))

    def test_main_offers_to_start_fresh_on_corrupt_settings(self):
        mainloop = mock.Mock()
        app = mock.Mock()
        app.user_data.corrupt = True
        for patcher in (
            mock.patch.object(tk, "Tk", lambda *_a, **_k: self.root),
            mock.patch.object(tk.Tk, "mainloop", mainloop),
            mock.patch.object(main, "StratagemApp", mock.Mock(return_value=app)),
            mock.patch.object(main.messagebox, "askyesno", return_value=True),
        ):
            patcher.start()
            self.addCleanup(patcher.stop)

        main.main()
        app.reset_corrupt_settings.assert_called_once()

    def test_sequence_for_unknown_name(self):
        app = self.build()
        self.assertEqual(app.sequence_for("Ghost Name"), "?")
        self.assertEqual(app.sequence_for("Alpha Strike"), "⬆ ⬅ ⬇ ➡")

    def test_set_stratagem_updates_widgets_and_saves(self):
        app = self.build()
        app.equipped[0] = "Belt Fed"
        app.set_stratagem(0, "Alpha Strike")

        self.assertEqual(app.equipped[0], "Alpha Strike")
        self.assertEqual(app.name_labels[0].cget("text"), "Alpha Strike")
        self.assertEqual(app.sequence_labels[0].cget("text"), "⬆ ⬅ ⬇ ➡")
        self.assertEqual(app.cooldown_labels[0].cget("text"), "Ready")
        self.assertEqual(app.cooldown_end[0], 0.0)
        self.assertEqual(self.saved()["equipped_stratagems"], app.equipped)

    def test_describe_ship_module_both_scopes(self):
        app = self.build()
        all_types = main.SHIP_MODULE_MAP["morale_augmentation"]
        partial = main.SHIP_MODULE_MAP["zero_g_breech_loading"]
        self.assertEqual(app.describe_ship_module(all_types), "-5% cooldown for all stratagems")
        self.assertEqual(app.describe_ship_module(partial), "-10% cooldown for Orbital stratagems")

    def test_ship_modules_panel_toggles_and_closes(self):
        app = self.build()
        app.open_ship_modules_panel()
        panel = next(w for w in self.root.winfo_children() if isinstance(w, tk.Toplevel))
        checks = [w for w in descendants(panel) if w.winfo_class() == "Checkbutton"]
        self.assertEqual(len(checks), len(main.SHIP_MODULES))

        for check in checks:
            check.invoke()
        self.assertEqual(app.active_ship_modules, set(main.SHIP_MODULE_MAP))
        self.assertEqual(self.saved()["ship_modules"], sorted(main.SHIP_MODULE_MAP))

        for check in checks:
            check.invoke()
        self.assertEqual(app.active_ship_modules, set())

        close = next(w for w in descendants(panel) if w.winfo_class() == "TButton")
        close.invoke()
        self.assertFalse(exists(panel))

    def test_icon_picker_closures_run_through_real_events(self):
        app = self.build()
        app.open_icon_picker(0)
        picker = next(w for w in self.root.winfo_children() if isinstance(w, tk.Toplevel))
        # Tk dispatches events only to mapped widgets, and a mapped widget needs a
        # mapped master. Alpha 0 keeps everything invisible while it is mapped.
        self.root.attributes("-alpha", 0)
        self.root.deiconify()
        picker.attributes("-alpha", 0)
        self.root.update()

        widgets = list(descendants(picker))
        canvas = next(w for w in widgets if w.winfo_class() == "Canvas")
        scrollbar = next(w for w in widgets if w.winfo_class() == "TScrollbar")
        entry = next(w for w in widgets if w.winfo_class() == "Entry")

        # rebuild_grid runs through the search variable write trace.
        entry.insert(0, "shield")
        entry.event_generate("<KeyRelease>")
        self.root.update()
        cells = [w for w in widgets if w.winfo_class() == "Frame" and w.grid_info()]
        self.assertEqual(len(cells), 2)

        # The scrollbar command is the nested on_scroll closure.
        # The scrollbar command is the nested on_scroll closure. In the real app it
        # runs when the pointer presses the trough; without a mainloop ttk does not
        # dispatch synthetic trough clicks, so call the registered Tcl command the
        # same way Tk does (shortcut: real trough clicks need a mainloop).
        app.last_picker_scroll = None
        app._enable_picker_scroll_capture()
        self.assertFalse(app.suppress_picker_scroll_capture)
        scrollbar.event_generate(
            "<Button-1>",
            x=scrollbar.winfo_width() // 2,
            y=scrollbar.winfo_height() // 2,
        )
        self.root.update()
        if app.last_picker_scroll is None:
            self.root.tk.call(scrollbar.cget("command"), "scroll", "1", "units")
        self.assertIsNotNone(app.last_picker_scroll)

        canvas.event_generate("<MouseWheel>", delta=-120)
        self.root.update()

        # A click on a grid cell selects the stratagem and closes the picker.
        cells[0].event_generate("<Button-1>")
        self.root.update()
        self.assertEqual(app.equipped[0], "Shield")
        self.assertFalse(exists(picker))

    def test_picker_scroll_capture_helpers(self):
        app = self.build()
        app.open_icon_picker(0)
        picker = next(w for w in self.root.winfo_children() if isinstance(w, tk.Toplevel))
        canvas = next(w for w in descendants(picker) if w.winfo_class() == "Canvas")

        app.suppress_picker_scroll_capture = True
        app._remember_picker_scroll(canvas)
        self.assertIsNone(app.last_picker_scroll)

        app._enable_picker_scroll_capture()
        app._remember_picker_scroll(canvas)
        self.assertEqual(app.last_picker_scroll, 0.0)

        # Reopening the picker restores the remembered scroll position.
        app.last_picker_scroll = 0.5
        app.pending_picker_restore = True
        app._restore_picker_scroll(canvas)
        self.assertFalse(app.pending_picker_restore)
        self.assertTrue(app.suppress_picker_scroll_capture)
        app._restore_picker_scroll_capture(False)
        self.assertFalse(app.suppress_picker_scroll_capture)
        picker.destroy()

    def test_wheel_handler_survives_a_destroyed_canvas(self):
        app = self.build()
        handlers = {}
        with mock.patch.object(tk.Misc, "bind_all", lambda _self, seq, func: handlers.setdefault(seq, func)):
            app.open_icon_picker(0)
            picker = next(w for w in self.root.winfo_children() if isinstance(w, tk.Toplevel))
            canvas = next(w for w in descendants(picker) if w.winfo_class() == "Canvas")
            canvas.destroy()
            handlers["<MouseWheel>"](SimpleNamespace(delta=-120))
            picker.destroy()

    def test_restore_picker_scroll_ignores_a_dead_canvas(self):
        app = self.build()
        app.open_icon_picker(0)
        picker = next(w for w in self.root.winfo_children() if isinstance(w, tk.Toplevel))
        canvas = next(w for w in descendants(picker) if w.winfo_class() == "Canvas")
        app.last_picker_scroll = 0.5
        app.pending_picker_restore = True

        moved = mock.Mock()
        canvas.yview_moveto = moved
        app._move_picker_scroll(canvas)
        self.assertEqual(moved.call_count, 1)

        with mock.patch.object(tk.Canvas, "winfo_exists", lambda _c: False):
            app._move_picker_scroll(canvas)
        self.assertEqual(moved.call_count, 1)
        picker.destroy()

    def test_presets(self):
        app = self.build()

        with mock.patch.object(main.simpledialog, "askstring", return_value=None) as ask:
            app.save_preset_prompt()
        ask.assert_called_once()
        self.assertEqual(app.presets, {})

        with mock.patch.object(main.simpledialog, "askstring", return_value="loadout A"):
            app.save_preset_prompt()
        self.assertEqual(app.presets, {"loadout A": app.equipped})
        self.assertEqual(app.active_preset, "loadout A")
        self.assertEqual(app.preset_combo.cget("values"), ("loadout A",))
        self.assertEqual(self.saved()["presets"], {"loadout A": app.equipped})

        app.save_preset_current()
        self.assertEqual(self.saved()["presets"], {"loadout A": app.equipped})

        with mock.patch.object(main.messagebox, "askyesno", return_value=False):
            app.delete_preset()
        self.assertIn("loadout A", app.presets)

        with mock.patch.object(main.messagebox, "askyesno", return_value=True):
            app.delete_preset()
        self.assertEqual(app.presets, {})
        self.assertEqual(app.active_preset, "")
        self.assertEqual(app.preset_var.get(), "")

        app.preset_var.set("")
        with mock.patch.object(main.messagebox, "askyesno", return_value=True):
            app.delete_preset()
        self.assertEqual(app.presets, {})

    def test_save_preset_current_without_active_preset_asks(self):
        app = self.build()
        self.assertEqual(app.active_preset, "")
        with mock.patch.object(main.simpledialog, "askstring", return_value="p") as ask:
            app.save_preset_current()
            ask.assert_called_once()
        self.assertEqual(app.active_preset, "p")

    def test_on_preset_select_branches(self):
        app = self.build()
        app.presets = {"loadout": ["Shield", "Belt Fed", "Alpha Strike"]}

        app.preset_var.set("")
        app.on_preset_select(None)
        self.assertEqual(app.active_preset, "")

        app.preset_var.set("missing")
        app.on_preset_select(None)
        self.assertEqual(app.active_preset, "")

        app.preset_var.set("loadout")
        app.on_preset_select(None)
        self.assertEqual(app.active_preset, "loadout")
        self.assertEqual(app.equipped[:3], ["Shield", "Belt Fed", "Alpha Strike"])
        self.assertEqual(app.name_labels[0].cget("text"), "Shield")
        self.assertEqual(self.saved()["active_preset"], "loadout")

    def test_on_key_delay_change(self):
        app = self.build()

        # IntVar.get() raises TclError on this Python, so the ValueError path needs a
        # stand-in variable whose get() returns a non-numeric string.
        real_var = app.key_delay_var
        app.key_delay_var = SimpleNamespace(get=lambda: "abc")
        app.on_key_delay_change()
        self.assertEqual(app.key_delay_ms, 40)
        app.key_delay_var = real_var

        app.key_delay_var.set(500)
        app.on_key_delay_change()
        self.assertEqual(app.key_delay_ms, 300)
        self.assertEqual(app.key_delay_var.get(), 300)
        self.assertEqual(self.saved()["key_delay_ms"], 300)

        app.key_delay_var.set(5)
        app.on_key_delay_change()
        self.assertEqual(app.key_delay_ms, 10)

    def test_on_input_keys_change(self):
        app = self.build()
        app.input_keys_var.set("Arrows")
        app.on_input_keys_change(None)
        self.assertEqual(app.input_keys, "arrows")
        app.input_keys_var.set("WASD")
        app.on_input_keys_change(None)
        self.assertEqual(app.input_keys, "wasd")
        self.assertEqual(self.saved()["input_keys"], "wasd")

    def test_register_hotkeys_success(self):
        app = self.build()
        self.assertEqual(len(app.hotkey_id_to_index), 9)
        self.assertEqual(app.hotkey_id_to_index[1000], 0)
        self.assertEqual(app.hotkey_id_to_index[1008], 8)
        self.assertEqual(app.status_var.get(), "Ready")

    def test_register_hotkeys_non_windows(self):
        app = self.build()
        with mock.patch.object(os, "name", "posix"):
            app.register_hotkeys()
        self.assertEqual(app.status_var.get(), "Hotkeys are supported only on Windows.")

    def test_register_hotkeys_runtime_error(self):
        app = self.build()
        with mock.patch.object(main, "HotkeyManager", fake_hotkey_manager(raise_runtime=True)):
            app.register_hotkeys()
        self.assertEqual(app.status_var.get(), "Hotkeys require Windows.")

    def test_register_hotkeys_reports_first_error(self):
        app = self.build()
        with mock.patch.object(main, "HotkeyManager", fake_hotkey_manager(errors=["boom"])):
            app.register_hotkeys()
        self.assertEqual(app.status_var.get(), "boom")

    def test_register_local_bindings(self):
        app = self.build()
        for keysym in main.LOCAL_KEYSYM_TO_INDEX:
            binding = app.root.bind(f"<KeyPress-{keysym}>")
            self.assertTrue(binding, keysym)
            self.assertIn("lambda", binding)

    def test_debug_key_capture_and_on_any_key(self):
        with mock.patch.object(main, "DEBUG_KEY_CAPTURE", True):
            app = self.build()
        self.assertTrue(app.root.bind_all("<KeyPress>"))

        app.on_any_key(SimpleNamespace())
        self.assertEqual(app.status_var.get(), "Key: keysym= keycode= char=")
        app.on_any_key(SimpleNamespace(keysym="A", keycode=65, char="a"))
        self.assertEqual(app.status_var.get(), "Key: keysym=A keycode=65 char=a")

    def test_hotkey_fired_and_ui_queue(self):
        app = self.build()
        app.on_hotkey_fired(9999)
        app.process_ui_queue()
        self.assertEqual(app.status_var.get(), "Ready")

        app.on_hotkey_fired(1000)
        app.process_ui_queue()
        self.assertEqual(app.status_var.get(), "Activated: Alpha Strike (W A S D)")
        self.assertEqual(self.user32.sent, [
            (0, 8),
            (0x57, 0), (0x57, 2),
            (0x41, 0), (0x41, 2),
            (0x53, 0), (0x53, 2),
            (0x44, 0), (0x44, 2),
            (0, 10),
        ])

        app.run_in_ui(lambda: (_ for _ in ()).throw(RuntimeError("bad")))
        app.process_ui_queue()
        self.assertEqual(app.status_var.get(), "UI update error: bad")

    def test_activate_stratagem(self):
        app = self.build()
        app.equipped[1] = "Ghost Name"
        app.activate_stratagem(1)
        self.assertEqual(app.status_var.get(), "Unknown stratagem: Ghost Name")

        with mock.patch.object(time, "monotonic", lambda: 1000.0):
            app.activate_stratagem(0)
            self.assertEqual(app.status_var.get(), "Activated: Alpha Strike (W A S D)")
            self.assertEqual(app.cooldown_end[0], 1120.0)

    def test_effective_cooldown(self):
        app = self.build()
        self.assertIsNone(app.effective_cooldown(app.stratagem_map["Belt Fed"]))

        strat = app.stratagem_map["Alpha Strike"]
        self.assertEqual(app.effective_cooldown(strat), 120)
        app.active_ship_modules = {"zero_g_breech_loading", "morale_augmentation"}
        self.assertAlmostEqual(app.effective_cooldown(strat), 120 * 0.85)

    def test_start_cooldown(self):
        app = self.build()
        with mock.patch.object(time, "monotonic", lambda: 1000.0):
            app.start_cooldown(1, app.stratagem_map["Belt Fed"])
            self.assertEqual(app.cooldown_end[1], 0.0)
            app.start_cooldown(0, app.stratagem_map["Alpha Strike"])
            self.assertEqual(app.cooldown_end[0], 1120.0)

    def test_format_cooldown(self):
        app = self.build()
        self.assertEqual(app.format_cooldown(70.5), "1:11")
        self.assertEqual(app.format_cooldown(3.2), "4s")
        # A whole second must not round up to the next one.
        self.assertEqual(app.format_cooldown(12.0), "12s")

    def test_update_cooldowns(self):
        app = self.build()
        app.equipped[1] = "Ghost Name"
        with mock.patch.object(time, "monotonic", lambda: 1000.0):
            app.cooldown_end[0] = 1100.0
            app.cooldown_end[1] = 1010.0
            app.cooldown_end[2] = 0.0
            app.update_cooldowns()
            self.assertEqual(app.cooldown_labels[0].cget("text"), "1:40")
            self.assertEqual(app.cooldown_labels[1].cget("text"), "10s")
            self.assertEqual(app.cooldown_labels[2].cget("text"), "Ready")

            app.cooldown_end[0] = 999.0
            app.update_cooldowns()
            self.assertEqual(app.cooldown_end[0], 0.0)
            self.assertEqual(app.cooldown_labels[0].cget("text"), "Ready")

    def test_send_sequence(self):
        app = self.build()
        self.user32.sent.clear()
        app.send_sequence(["W", "Q"])
        self.assertEqual(self.user32.sent, [(0, 8), (0x57, 0), (0x57, 2), (0, 10)])

        app.input_keys = "arrows"
        self.user32.sent.clear()
        app.send_sequence(["W"])
        self.assertEqual(self.user32.sent, [(0, 8), (0x26, 0), (0x26, 2), (0, 10)])

    def test_send_sequence_returns_off_windows(self):
        app = self.build()
        self.user32.sent.clear()
        with mock.patch.object(os, "name", "posix"):
            app.send_sequence(["W"])
        self.assertEqual(self.user32.sent, [])

    def test_blocked_sendinput_still_releases_ctrl(self):
        app = self.build()
        self.user32.sent.clear()
        self.user32.failing = {0}
        app.send_sequence(["W"])

        self.assertEqual(self.user32.sent, [(0, 8), (0x57, 0), (0x57, 2), (0, 10)])
        self.assertEqual(
            app.status_var.get(),
            "SendInput blocked - run this app as administrator to type into an elevated game.",
        )
        self.assertFalse(app.sending)

    def test_blocked_ctrl_press_reports_and_stops(self):
        app = self.build()
        self.user32.sent.clear()
        self.user32.failing = {8}
        app.send_sequence(["W"])

        self.assertEqual(self.user32.sent, [(0, 8)])
        self.assertEqual(
            app.status_var.get(),
            "SendInput blocked - run this app as administrator to type into an elevated game.",
        )
        self.assertFalse(app.sending)

    def test_ctrl_released_when_a_send_raises(self):
        app = self.build()
        self.user32.sent.clear()
        self.user32.raising = {0}
        with self.assertRaises(RuntimeError):
            app.send_sequence(["W"])

        self.assertEqual(self.user32.sent, [(0, 8), (0x57, 0), (0, 10)])
        self.assertFalse(app.sending)

    def test_second_activation_while_typing_is_refused(self):
        app = self.build()
        self.user32.sent.clear()
        app.sending = True
        app.activate_stratagem(0)
        self.assertEqual(app.status_var.get(), "Still typing the previous sequence.")
        self.assertEqual(self.user32.sent, [])

    def test_on_close_stops_hotkeys(self):
        app = self.build()
        stopped = []
        app.hotkeys = SimpleNamespace(stop=lambda: stopped.append("stop"))
        app.on_close()
        self.assertEqual(stopped, ["stop"])
        self.assertFalse(exists(app.root))

    def test_on_close_without_hotkeys(self):
        app = self.build()
        app.hotkeys = None
        app.on_close()
        self.assertFalse(exists(app.root))


class MainEntryPointTest(unittest.TestCase):
    def setUp(self):
        tmp = tempfile.TemporaryDirectory()
        self.addCleanup(tmp.cleanup)
        patcher = mock.patch.dict(os.environ, {"APPDATA": tmp.name})
        patcher.start()
        self.addCleanup(patcher.stop)
        # Self-rescheduling after() scripts would fire after the root is destroyed.
        patcher = mock.patch.object(tk.Misc, "after", lambda *_a, **_k: None)
        patcher.start()
        self.addCleanup(patcher.stop)

        self.root = tk.Tk()
        self.root.withdraw()
        self.addCleanup(lambda: exists(self.root) and self.root.destroy())

    def run_as_main(self):
        saved = sys.modules.pop("main", None)
        try:
            return runpy.run_module("main", run_name="__main__")
        finally:
            if saved is not None:
                sys.modules["main"] = saved

    def test_main_runs_and_the_module_guard_executes(self):
        mainloop = mock.Mock()
        for patcher in (
            mock.patch.object(tk, "Tk", lambda *_a, **_k: self.root),
            mock.patch.object(tk.Tk, "mainloop", mainloop),
            mock.patch.object(ctypes, "windll", fake_windll(FakeUser32())),
            mock.patch.object(ctypes, "byref", lambda obj: obj),
        ):
            patcher.start()
            self.addCleanup(patcher.stop)

        namespace = self.run_as_main()
        self.assertEqual(namespace["__name__"], "__main__")
        mainloop.assert_called_once()
        # main() built the app on the root we handed it.
        self.assertTrue(self.root.winfo_children())

    def test_main_reports_startup_errors(self):
        mainloop = mock.Mock()
        showerror = mock.Mock()
        for patcher in (
            mock.patch.object(tk, "Tk", lambda *_a, **_k: self.root),
            mock.patch.object(tk.Tk, "mainloop", mainloop),
            mock.patch.object(main, "StratagemApp", mock.Mock(side_effect=RuntimeError("bad icon"))),
            mock.patch.object(main.messagebox, "showerror", showerror),
        ):
            patcher.start()
            self.addCleanup(patcher.stop)

        main.main()
        showerror.assert_called_once_with("pystrat", "Failed to start: bad icon")
        mainloop.assert_not_called()

    def test_main_survives_a_theme_without_tcl_error_support(self):
        mainloop = mock.Mock()
        theme_use = mock.Mock(side_effect=tk.TclError)
        for patcher in (
            mock.patch.object(tk, "Tk", lambda *_a, **_k: self.root),
            mock.patch.object(tk.Tk, "mainloop", mainloop),
            mock.patch.object(ttk.Style, "theme_use", theme_use),
            mock.patch.object(ctypes, "windll", fake_windll(FakeUser32())),
            mock.patch.object(ctypes, "byref", lambda obj: obj),
        ):
            patcher.start()
            self.addCleanup(patcher.stop)

        namespace = self.run_as_main()
        self.assertEqual(namespace["__name__"], "__main__")
        theme_use.assert_called_once_with("clam")
        mainloop.assert_called_once()
        self.assertTrue(self.root.winfo_children())

