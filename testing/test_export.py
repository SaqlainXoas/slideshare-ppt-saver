from __future__ import annotations

import io
import sys
import tempfile
import unittest
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))

from PIL import Image
from pptx import Presentation

from slideshare_to_pptx import ExportResult, export_slideshare_pptx


def build_webp_bytes(color: tuple[int, int, int]) -> bytes:
    buffer = io.BytesIO()
    image = Image.new("RGB", (1280, 720), color=color)
    image.save(buffer, format="WEBP")
    return buffer.getvalue()


class FakeResponse:
    def __init__(
        self,
        *,
        status_code: int = 200,
        text: str = "",
        content: bytes = b"",
        headers: dict[str, str] | None = None,
        url: str = "https://www.slideshare.net/slideshow/demo/123",
    ) -> None:
        self.status_code = status_code
        self.text = text
        self.content = content
        self.headers = headers or {}
        self.url = url

    @property
    def ok(self) -> bool:
        return 200 <= self.status_code < 300

    def raise_for_status(self) -> None:
        if not self.ok:
            raise RuntimeError(f"HTTP {self.status_code}")


class FakeSession:
    def __init__(self, fixture_html: str) -> None:
        self.fixture_html = fixture_html
        self.slide_bytes = {
            1: build_webp_bytes((220, 38, 38)),
            2: build_webp_bytes((37, 99, 235)),
            3: build_webp_bytes((22, 163, 74)),
        }

    def get(self, url: str, timeout: int) -> FakeResponse:
        if "slideshare.net" in url:
            return FakeResponse(
                text=self.fixture_html,
                headers={"content-type": "text/html; charset=utf-8"},
                url=url,
            )

        if "Clinical-Deck-1-2048" in url:
            return FakeResponse(
                content=self.slide_bytes[1],
                headers={"content-type": "image/webp"},
                url=url,
            )
        if "Clinical-Deck-2-1024" in url:
            return FakeResponse(
                content=self.slide_bytes[2],
                headers={"content-type": "image/webp"},
                url=url,
            )
        if "Clinical-Deck-3-2048" in url:
            return FakeResponse(
                content=self.slide_bytes[3],
                headers={"content-type": "image/webp"},
                url=url,
            )

        return FakeResponse(status_code=404, headers={"content-type": "text/plain"}, url=url)


class ExportIntegrationTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls) -> None:
        fixture_path = ROOT / "testing" / "fixtures" / "sample_slideshare_page.html"
        cls.fixture_html = fixture_path.read_text(encoding="utf-8")

    def test_export_slideshare_pptx_builds_expected_presentation(self) -> None:
        session = FakeSession(self.fixture_html)
        with tempfile.TemporaryDirectory() as temp_dir:
            result = export_slideshare_pptx(
                url="https://www.slideshare.net/slideshow/clinical-deck/123",
                output_dir=temp_dir,
                session=session,
            )

            self.assertIsInstance(result, ExportResult)
            self.assertEqual(result.title, "Clinical Deck")
            self.assertEqual(result.slide_count, 3)
            self.assertEqual(result.pptx_path.name, "Clinical Deck.pptx")
            self.assertTrue(result.pptx_path.exists())

            presentation = Presentation(result.pptx_path)
            self.assertEqual(len(presentation.slides), 3)

            image_files = sorted((result.pptx_path.parent / "images").glob("slide-*.png"))
            self.assertEqual(len(image_files), 3)
            self.assertTrue((result.pptx_path.parent / "SOURCE.txt").exists())


if __name__ == "__main__":
    unittest.main()
