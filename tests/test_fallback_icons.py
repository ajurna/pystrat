import json
import tempfile
import unittest
from pathlib import Path
from unittest import mock

import main

# Stratagem types that have a dedicated fallback icon. "resupply" has none:
# every resupply stratagem is expected to ship its own icon.
FALLBACK_ICONS = {
    "orbital": "Orbital Fallback.svg",
    "eagle": "Eagle Fallback.svg",
    "mission": "Mission Fallback.svg",
    "backpack": "Backpack Fallback.svg",
    "support_weapon": "Support Weapon Fallback.svg",
    "sentry": "Sentry Fallback.svg",
    "tank": "Tank Fallback.svg",
    "vehicle": "Vehicle Fallback.svg",
    "exosuit": "Exosuit Fallback.svg",
    "emplacement": "Emplacement Fallback.svg",
}

PNG_SIGNATURE = b"\x89PNG\r\n\x1a\n"


class FallbackIconFilesTest(unittest.TestCase):
    def test_fallback_file_exists_for_each_type(self):
        for stratagem_type, filename in FALLBACK_ICONS.items():
            with self.subTest(type=stratagem_type):
                self.assertTrue((main.ICON_DIR / filename).is_file())

    def test_fallback_types_are_known_types(self):
        self.assertLessEqual(FALLBACK_ICONS.keys(), main.ALL_STRATAGEM_TYPES)

    def test_every_fallback_icon_has_a_test_entry(self):
        on_disk = {path.name for path in main.ICON_DIR.glob("* Fallback.svg")}
        self.assertEqual(on_disk, set(FALLBACK_ICONS.values()))

    def test_fallback_icons_render_to_png(self):
        for stratagem_type, filename in FALLBACK_ICONS.items():
            with self.subTest(type=stratagem_type):
                png = main.render_svg_to_png_bytes(main.ICON_DIR / filename, 64)
                self.assertIsNotNone(png)
                self.assertTrue(png.startswith(PNG_SIGNATURE))


class FallbackIconSelectionTest(unittest.TestCase):
    """Drives the real load_stratagems against a temporary icon directory."""

    def setUp(self):
        tmp = tempfile.TemporaryDirectory()
        self.addCleanup(tmp.cleanup)
        self.tmp = Path(tmp.name)
        self.icon_dir = self.tmp / "icons"
        self.icon_dir.mkdir()
        self.stratagems_file = self.tmp / "stratagems.json"

        self.rendered: list[Path] = []

        def fake_render(svg_path, size):
            self.rendered.append(Path(svg_path))
            return b"png"

        for patcher in (
            mock.patch.object(main, "ICON_DIR", self.icon_dir),
            mock.patch.object(main, "STRATAGEMS_FILE", self.stratagems_file),
            mock.patch.object(main, "render_svg_to_png_bytes", fake_render),
            mock.patch.object(main.tk, "PhotoImage"),
        ):
            patcher.start()
            self.addCleanup(patcher.stop)

    def load(self, entries):
        self.stratagems_file.write_text(json.dumps(entries), encoding="utf-8")
        return main.load_stratagems()

    @staticmethod
    def entry(name, stratagem_type):
        return {"name": name, "sequence": ["W"], "type": stratagem_type}

    def test_missing_icon_uses_type_fallback(self):
        for stratagem_type, filename in FALLBACK_ICONS.items():
            with self.subTest(type=stratagem_type):
                self.rendered.clear()
                self.load([self.entry("No Icon", stratagem_type)])
                self.assertEqual(self.rendered, [self.icon_dir / filename])

    def test_own_icon_is_preferred_over_fallback(self):
        (self.icon_dir / "Has Icon.svg").write_text("<svg/>", encoding="utf-8")
        self.load([self.entry("Has Icon", "orbital")])
        self.assertEqual(self.rendered, [self.icon_dir / "Has Icon.svg"])

    def test_fallback_is_chosen_per_entry(self):
        (self.icon_dir / "Has Icon.svg").write_text("<svg/>", encoding="utf-8")
        self.load(
            [
                self.entry("No Icon A", "eagle"),
                self.entry("Has Icon", "eagle"),
                self.entry("No Icon B", "tank"),
            ]
        )
        self.assertEqual(
            self.rendered,
            [
                self.icon_dir / "Eagle Fallback.svg",
                self.icon_dir / "Has Icon.svg",
                self.icon_dir / "Tank Fallback.svg",
            ],
        )


class BundledStratagemIconsTest(unittest.TestCase):
    def test_every_stratagem_has_an_icon_or_a_fallback(self):
        raw = json.loads(main.STRATAGEMS_FILE.read_text(encoding="utf-8"))
        for entry in raw:
            with self.subTest(name=entry["name"]):
                own = main.ICON_DIR / f"{entry['name']}.svg"
                fallback_name = FALLBACK_ICONS.get(entry["type"])
                fallback = main.ICON_DIR / fallback_name if fallback_name else None
                self.assertTrue(
                    own.is_file() or (fallback is not None and fallback.is_file()),
                    f"no icon or fallback for {entry['name']} ({entry['type']})",
                )


if __name__ == "__main__":
    unittest.main()
