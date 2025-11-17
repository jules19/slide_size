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

# Generate optimization recommendations
python pptx_heavy_slides.py presentation.pptx --optimization-report
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
- `--optimization-report`: Generate optimization recommendations (conference-quality focused)
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

## Optimization Analysis (NEW FEATURE)

### Purpose
Analyzes presentations for image optimization opportunities while maintaining conference-quality standards. Designed for presentations projected on large screens (1920x1080 Full HD typical).

### Detection Logic (Conservative Thresholds)

**1. Oversized Resolution** (>2.5x display size)
- Compares actual image pixel dimensions vs. display size on slide
- Threshold: Image > 2.5x display dimensions
- Recommendation: Resize to 2x display size (retina quality)
- Priority: HIGH if >5x, MEDIUM if 2.5-5x

**2. Absolute Size Caps** (>3200px longest edge)
- Safety net for unreasonably large images
- Recommendation: Resize to 2560px max (suitable for conference projectors)
- Priority: MEDIUM

**3. PNG Photos** (PNG >1MB)
- Detects large PNG files that should be JPEG
- Recommendation: Convert to JPEG quality 85-90
- Priority: MEDIUM if >3MB, LOW otherwise

**4. Uncompressed JPEG** (>1 byte/pixel)
- Detects quality 95-100 JPEG that could use lower quality
- Recommendation: Re-save at quality 85
- Priority: LOW

### Data Model

```python
class ImageDimensions(TypedDict):
    pixel_width: int
    pixel_height: int
    display_width_px: int
    display_height_px: int
    resolution_ratio: float  # How many times larger than display

class OptimizationOpportunity(TypedDict):
    slide_index: int
    slide_title: str | None
    opportunity_type: str  # "oversized_resolution", "absolute_size", "png_photo", "uncompressed_jpeg"
    current_bytes: int
    potential_bytes: int
    savings_bytes: int
    savings_percent: float
    current_dimensions: str
    display_dimensions: str
    recommended_dimensions: str
    current_format: str
    recommended_format: str
    details: str
    severity: str  # "high", "medium", "low"
    is_shared: bool
```

### Key Functions

- `get_image_dimensions(image_blob, shape)`: Extracts pixel and display dimensions
- `analyze_image_optimization(...)`: Analyzes single image for opportunities
- `analyze_optimization_opportunities(path)`: Main analysis function
- `print_optimization_report(...)`: Formatted console output

### Implementation Notes

- Uses Pillow to extract actual pixel dimensions from image blob
- Converts shape dimensions from EMUs to pixels (96 DPI standard)
- Detects shared media (optimization affects all slides using image)
- Conservative thresholds prioritize visual quality for conference projection
- Estimates savings based on pixel reduction ratios

## Future Enhancements

Potential improvements noted in code:
- **Strategy B**: ZIP/relationships parsing for more precise media extraction
- **Enhanced media detection**: Better video/audio relationship handling
- **Relationship IDs**: Currently set to None, could be extracted from PPTX relationships
- **Actual Optimization**: Implement blob replacement or post-processing to apply optimizations
- **Dry-run mode**: Show optimization preview before applying changes
- **Batch processing**: Optimize multiple presentations at once
