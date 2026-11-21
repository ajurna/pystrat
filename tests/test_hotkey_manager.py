import os
import unittest
from types import SimpleNamespace
from unittest import mock

import main


def fake_user32(get_message, register_ok=True):
    calls = []
    return SimpleNamespace(
        calls=calls,
        PeekMessageW=lambda *_a: 0,
        RegisterHotKey=lambda *_a: register_ok,
        UnregisterHotKey=lambda *_a: 1,
        GetMessageW=get_message,
        PostThreadMessageW=lambda *_a: calls.append("quit"),
    )


def get_message_factory(events):
    pending = list(events)

    def get_message(msg, *_args):
        if not pending:
            return 0
        msg.message = main.WM_HOTKEY
        msg.wParam = pending.pop(0)
        return 1

    return get_message


class HotkeyManagerTest(unittest.TestCase):
    def make(self, notify, events=(), register_ok=True):
        with mock.patch.object(os, "name", "nt"):
            manager = main.HotkeyManager(notify)

        manager.user32 = fake_user32(get_message_factory(events), register_ok)
        manager.ctypes = SimpleNamespace(
            windll=SimpleNamespace(
                kernel32=SimpleNamespace(GetCurrentThreadId=lambda: 4242)
            ),
            byref=lambda obj: obj,
        )
        return manager

    def test_requires_windows(self):
        with mock.patch.object(os, "name", "posix"), self.assertRaises(RuntimeError) as ctx:
                main.HotkeyManager(lambda _i: None)
        self.assertEqual(str(ctx.exception), "Hotkeys require Windows.")

    def test_stop_before_start_does_nothing(self):
        manager = self.make(lambda _i: None)
        self.assertIsNone(manager.thread_id)
        self.assertIsNone(manager.thread)
        manager.stop()

    def test_message_loop_notifies_then_unregisters(self):
        fired = []
        manager = self.make(fired.append, events=[1000, 1001])
        manager.start({1000: 0x57, 1001: 0x41})

        manager.thread.join(5)
        self.assertEqual(fired, [1000, 1001])
        self.assertEqual(manager.errors, [])
        self.assertEqual(manager.thread_id, 4242)
        self.assertEqual(sorted(manager.hotkey_map), [1000, 1001])

        manager.stop()

    def test_get_message_error_records_error_and_stops(self):
        manager = self.make(lambda _i: None)
        manager.user32.GetMessageW = lambda *_a: -1
        manager.start({1000: 0x57})
        manager.thread.join(5)

        self.assertEqual(manager.errors, ["Hotkey message loop stopped (GetMessageW failed)."])
        manager.stop()

    def test_callback_failure_is_recorded_and_loop_continues(self):
        def notify(hotkey_id):
            raise RuntimeError("bad callback")

        manager = self.make(notify, events=[1000])
        manager.start({1000: 0x57})
        manager.thread.join(5)

        self.assertEqual(manager.errors, ["Hotkey callback failed: bad callback"])
        manager.stop()

    def test_register_failure_records_error(self):
        fired = []
        manager = self.make(fired.append, events=[], register_ok=False)
        manager.start({1000: 0x57})
        manager.thread.join(5)

        self.assertEqual(manager.errors, ["Failed to register hotkey 1000"])
        self.assertEqual(fired, [])
        manager.stop()

