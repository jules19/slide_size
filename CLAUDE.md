# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

PowerPoint "Heavy Slides" Analyzer - A Python CLI tool that analyzes .pptx files to identify which slides contribute most to file size due to embedded media (images, videos, audio).

**Target**: Python 3.10+, cross-platform (Linux, macOS, Windows)

**Dependencies**: `python-pptx`, `pytest`, `Pillow` (see requirements.txt)

## Architecture

### Main Module: pptx_heavy_slides.py

**Core function** (reusable as library):
```python
def analyze_pptx_media(path: str, include_shared_media: bool = False) -> list[SlideMediaStats]
```

**Key implementation details**:
- Uses `python-pptx` library with Strategy A (shape objects, not ZIP/XML parsing)
- **Shared media detection**: Uses hash of `image.blob` as key in `media_registry` dictionary
  - Tracks which slides use each media item
  - Default behavior (`--ignore-shared-media`): counts bytes only on first slide appearance
  - With `--include-shared-media`: counts bytes on every slide
- **Two-pass algorithm**:
  1. First pass: Collect all media, build registry, detect sharing
  2. Second pass: Build SlideMediaStats with proper byte counting
- Handles images via `MSO_SHAPE_TYPE.PICTURE` and `shape.image.blob`
- Video/audio detection via `MSO_SHAPE_TYPE.MEDIA` (note: may need enhancement for complex media)

**Helper functions**:
- `get_slide_title()`: Extracts title from slide.shapes.title
- `format_bytes()`: Converts bytes to human-readable (B/KB/MB/GB)
- `print_console_output()`: Human-readable ranked output
- `write_json_output()` / `write_csv_output()`: Export functions
- `main()`: CLI entry point with argparse

### Data Model

```python
class SlideMediaStats(TypedDict):
    slide_index: int            # 1-based
    slide_title: str | None
    total_media_bytes: int
    image_bytes: int
    video_bytes: int
    audio_bytes: int
    other_media_bytes: int
    media_items: list[MediaItem]

class MediaItem(TypedDict):
    type: str                   # "image", "video", "audio", "other"
    size_bytes: int
    filename: str | None
    content_type: str | None
    relationship_id: str | None
    shared: bool                # True if appears on >1 slide
```

## Development Commands

### Installation
```bash
pip install -r requirements.txt
```

### Running the Tool
```bash
# Basic analysis
python pptx_heavy_slides.py presentation.pptx

# Top 5 slides with verbose output
python pptx_heavy_slides.py presentation.pptx --top 5 --verbose

# Export to JSON and CSV
python pptx_heavy_slides.py presentation.pptx --output-json results.json --output-csv results.csv

# Include shared media in counts (count on every slide)
python pptx_heavy_slides.py presentation.pptx --include-shared-media
```

### Testing
```bash
# Run all tests
pytest

# Run with verbose
pytest -v

# Run specific test
pytest tests/test_analyzer.py::test_shared_media_ignore_mode

# Generate sample presentation for manual testing
python create_sample.py  # Creates sample_presentation.pptx
```

### Test Coverage

The test suite (`tests/test_analyzer.py`) includes:
- Single-slide and multi-slide presentations
- Shared media detection (both `--ignore` and `--include` modes)
- Empty decks (no media)
- Error handling (file not found, wrong extension, corrupt PPTX)
- JSON/CSV output validation
- Byte formatting
- Uses Pillow to generate synthetic test images

## CLI Arguments

- `input_path` (positional): Path to .pptx file
- `--top N`: Show only top N heaviest slides
- `--output-json <path>`: Write results as JSON
- `--output-csv <path>`: Write results as CSV
- `--include-shared-media`: Count shared media on every slide
- `--ignore-shared-media`: Count shared media only once (default)
- `--verbose`: Enable debug logging
- `--version`: Show version and exit

## Error Handling

**Exit codes**: 0 (success), 1 (all errors)

**Error messages** (to stderr):
- File not found: `Error: file not found: <path>`
- Wrong type: `Error: unsupported file type (expected .pptx): <path>`
- Corrupt: `Error: failed to open .pptx: <details>`
- Output write failure: `Error: failed to write output file <path>: <details>`

Uses Python's `logging` module for internal logging (not print statements).

## Key Behaviors

**Shared media example**: If same image appears on slides 3 and 4:
- Default (`--ignore-shared-media`): Slide 3 shows full bytes, Slide 4 shows 0 bytes
- With `--include-shared-media`: Both slides show full bytes

**Slide titles**: Extracted from `slide.shapes.title.text`. Displays "(no title)" if unavailable.

**Output sorting**: Always sorted by `total_media_bytes` descending (heaviest first).

## Future Enhancements

Potential improvements noted in code:
- **Strategy B**: ZIP/relationships parsing for more precise media extraction
- **Enhanced media detection**: Better video/audio relationship handling
- **Relationship IDs**: Currently set to None, could be extracted from PPTX relationships
