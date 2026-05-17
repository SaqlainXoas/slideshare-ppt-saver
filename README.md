# SlideShare PPT Saver

Turn a public SlideShare deck into a downloadable `.pptx` with the detected deck title as the filename.

## Features

- extracts slide images from public SlideShare pages
- prefers higher-resolution images when available
- builds a `.pptx` automatically
- uses the detected deck title for the output filename
- includes a simple Streamlit UI for paste-and-download use
- includes `pytest` coverage for core extraction, export, and UI behavior

## Important Limitation

This tool creates an image-based PowerPoint.

It does **not** recover the original fully editable source deck with separate text boxes, charts, or shapes.

Use it only for decks you are allowed to archive, reuse, or download.

## Requirements

- Python `3.12`
- [`uv`](https://docs.astral.sh/uv/)

## Quick Start

```bash
uv sync
uv run streamlit run app.py
```

Then open the local URL, paste a public SlideShare link, and click `Generate`.

## CLI Usage

```bash
uv run python slideshare_to_pptx.py "https://www.slideshare.net/slideshow/example/123"
```

Generated files are written to `output/<deck title>/`.

## Testing

```bash
uv run pytest
```

## Deploy

This project is a good fit for free hosting on Streamlit Community Cloud or Hugging Face Spaces.

## Project Files

- `app.py` - Streamlit entrypoint
- `ui.py` - UI layout and interaction flow
- `slideshare_to_pptx.py` - SlideShare extraction and PPTX export logic
- `testing/` - automated tests

## Notes

- Output naming is based on the detected SlideShare title.
- Public SlideShare page changes can affect extraction behavior over time.
