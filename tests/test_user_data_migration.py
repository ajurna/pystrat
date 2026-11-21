import json
import tempfile
import unittest
from pathlib import Path

import main


class UserDataMigrationTest(unittest.TestCase):
    def load(self, payload):
        with tempfile.TemporaryDirectory() as tmp:
            path = Path(tmp) / "user_data.json"
            path.write_text(json.dumps(payload), encoding="utf-8")
            return main.UserData.load(path)

    def test_equipped_and_presets_use_current_names(self):
        data = self.load(
            {
                "equipped_stratagems": ["CQC-20", "Eagle Airstrike"],
                "presets": {"loadout A": ["CQC-20", "Reinforce"]},
            }
        )
        self.assertEqual(data.equipped_stratagems, ["Breaching Hammer", "Eagle Airstrike"])
        self.assertEqual(data.presets, {"loadout A": ["Breaching Hammer", "Reinforce"]})

    def test_names_not_in_legacy_map_are_untouched(self):
        data = self.load({"equipped_stratagems": ["Orbital Laser"], "presets": {"p": ["Solo Silo"]}})
        self.assertEqual(data.equipped_stratagems, ["Orbital Laser"])
        self.assertEqual(data.presets, {"p": ["Solo Silo"]})

    def test_corrupt_file_falls_back_to_defaults(self):
        with tempfile.TemporaryDirectory() as tmp:
            path = Path(tmp) / "user_data.json"
            path.write_text("{ not json", encoding="utf-8")
            data = main.UserData.load(path)

        self.assertTrue(data.corrupt)
        self.assertEqual(data.equipped_stratagems, [])
        self.assertEqual(data.key_delay_ms, 40)
        self.assertEqual(data.input_keys, "wasd")

    def test_unknown_input_keys_fall_back_to_wasd(self):
        data = self.load({"equipped_stratagems": ["Orbital Laser"], "input_keys": "vim"})
        self.assertEqual(data.input_keys, "wasd")
