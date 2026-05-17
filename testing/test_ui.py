from __future__ import annotations

import sys
import unittest
from pathlib import Path
from unittest.mock import patch

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))

from streamlit.testing.v1 import AppTest


class UIBehaviourTests(unittest.TestCase):
    def test_invalid_url_shows_validation_error(self) -> None:
        app = AppTest.from_file(str(ROOT / "app.py"))
        app.run()
        app.text_input[0].set_value("https://example.com/nope").run()
        app.button[0].click().run()

        self.assertEqual(len(app.error), 1)
        self.assertIn("Only public SlideShare URLs are supported", app.error[0].value)

    def test_success_path_renders_download_state(self) -> None:
        with patch("ui.build_download_payload", return_value=(b"pptx", "Deck.pptx", 3, "Deck")):
            app = AppTest.from_file(str(ROOT / "app.py"))
            app.run()
            app.text_input[0].set_value(
                "https://www.slideshare.net/slideshow/deck/123"
            ).run()
            app.button[0].click().run()

            self.assertEqual(len(app.error), 0)
            self.assertTrue(any("Deck" in item.value for item in app.markdown))
            self.assertTrue(any("3 slides" in item.value for item in app.markdown))


if __name__ == "__main__":
    unittest.main()
