from __future__ import annotations

import sys
import unittest
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))

from slideshare_to_pptx import (
    SlideShareError,
    build_download_candidates,
    extract_slide_urls,
    extract_title,
    normalize_slideshare_url,
    sanitize_filename,
)


class CoreExtractionTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls) -> None:
        fixture_path = ROOT / "testing" / "fixtures" / "sample_slideshare_page.html"
        cls.fixture_html = fixture_path.read_text(encoding="utf-8")

    def test_normalize_slideshare_url_adds_https(self) -> None:
        url = normalize_slideshare_url("www.slideshare.net/slideshow/demo/123")
        self.assertEqual(url, "https://www.slideshare.net/slideshow/demo/123")

    def test_normalize_slideshare_url_rejects_other_sites(self) -> None:
        with self.assertRaises(SlideShareError):
            normalize_slideshare_url("https://example.com/demo")

    def test_sanitize_filename_removes_invalid_characters(self) -> None:
        self.assertEqual(
            sanitize_filename('Clinical: Deck / Final?.pptx'),
            "Clinical Deck Final.pptx",
        )

    def test_extract_title_uses_clean_deck_name(self) -> None:
        self.assertEqual(extract_title(self.fixture_html), "Clinical Deck")

    def test_extract_slide_urls_keeps_best_resolution_per_slide(self) -> None:
        urls = extract_slide_urls(self.fixture_html)
        self.assertEqual(len(urls), 3)
        self.assertTrue(urls[1].endswith("Clinical-Deck-1-2048.jpg"))
        self.assertTrue(urls[2].endswith("Clinical-Deck-2-638.jpg"))
        self.assertTrue(urls[3].endswith("Clinical-Deck-3-1600.jpg"))

    def test_extract_slide_urls_rejects_missing_slide_numbers(self) -> None:
        html = """
        <img src="https://image.slidesharecdn.com/demo/85/Demo-1-320.jpg" />
        <img src="https://image.slidesharecdn.com/demo/85/Demo-3-320.jpg" />
        """
        with self.assertRaises(SlideShareError):
            extract_slide_urls(html)

    def test_build_download_candidates_prefers_higher_resolutions(self) -> None:
        candidates = build_download_candidates(
            "https://image.slidesharecdn.com/demo/85/Demo-2-320.jpg"
        )
        self.assertTrue(candidates[0].endswith("Demo-2-2048.jpg"))
        self.assertIn(
            "https://image.slidesharecdn.com/demo/85/Demo-2-320.jpg",
            candidates,
        )


if __name__ == "__main__":
    unittest.main()
