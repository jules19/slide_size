# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

PowerPoint "Heavy Slides" Analyzer - A Python CLI tool that analyzes .pptx files to identify which slides contribute most to file size due to embedded media (images, videos, audio).

**Target**: Python 3.10+, cross-platform (Linux, macOS, Windows)

**Primary dependency**: `python-pptx` (or standard library zipfile + XML parsing)

## Key Functional Requirements

### CLI Interface

Main entry point: `python pptx_heavy_slides.py <path-to-pptx> [options]`

**Critical CLI arguments**:
- `input_path` (positional): Path to .pptx file
- `--top N`: Show only top N heaviest slides
- `--output-json <path>`: Write results as JSON
- `--output-csv <path>`: Write results as CSV
- `--include-shared-media` / `--ignore-shared-media`: Handle shared media counting
  - **Default is `--ignore-shared-media`**: If an image/media is used on multiple slides, count bytes only once on the first slide it appears
  - `--include-shared-media`: Count the media bytes on every slide that uses it
- `--verbose`: Debug logging
- `--version`: Tool version

### Core Data Model

```python
class SlideMediaStats(TypedDict):
    slide_index: int            # 1-based index
    slide_title: str | None
    total_media_bytes: int
    image_bytes: int
    video_bytes: int
    audio_bytes: int
    other_media_bytes: int
    media_items: list[dict]     # each with type, size_bytes, filename, etc.
```

### Architecture Guidelines

**Separation of concerns**:
- Core analysis logic must be separate from CLI parsing
- Main analysis function signature:
  ```python
  def analyze_pptx_media(path: str, include_shared_media: bool = False) -> list[SlideMediaStats]:
      ...
  ```
- CLI entry point under `if __name__ == "__main__":`

**Shared media detection**:
- Create global dictionary keyed by media identifier (filename/partname in PPTX package)
- Track which slides use each media file
- Mark items as `shared: True` if used on >1 slide
- Apply --ignore-shared-media logic: only count bytes on first slide appearance

**Implementation strategy**:
- Strategy A (preferred for v1): Use `python-pptx` shape objects
  - Iterate `slide.shapes`, check `shape.shape_type` for PICTURE
  - Extract bytes via `shape.image.blob`
  - Handle media via `shape.media_format` or relationships
- Strategy B (future enhancement): ZIP/relationships parsing
  - Parse slideN.xml for `<a:blip>` references
  - Map relationship IDs to /ppt/media/* files
  - Use ZipFile.getinfo() for exact sizes

## Development Commands

**Testing** (once implemented):
```bash
# Run all tests
pytest

# Run specific test
pytest tests/test_analyzer.py::test_shared_media

# Run with verbose
pytest -v
```

**Running the tool**:
```bash
# Basic analysis
python pptx_heavy_slides.py presentation.pptx

# Top 5 slides with verbose output
python pptx_heavy_slides.py presentation.pptx --top 5 --verbose

# Export to JSON and CSV
python pptx_heavy_slides.py presentation.pptx --output-json results.json --output-csv results.csv

# Include shared media in counts
python pptx_heavy_slides.py presentation.pptx --include-shared-media
```

## Testing Requirements

Must include tests for:
1. Single-slide deck with one image (verify non-zero bytes)
2. Multi-slide deck with same image on multiple slides (test both --ignore-shared-media and --include-shared-media)
3. Deck with video and audio (verify categorization)
4. Empty deck (all slides should have 0 bytes)
5. Invalid inputs (non-existent file, wrong extension, corrupt PPTX → errors + non-zero exit codes)

Use pytest or unittest. Generate test PPTX files using python-pptx in test setup.

## Error Handling

**Exit codes**:
- 0: Success
- 1: All errors (file not found, invalid PPTX, corrupt file, output write failures)

**Error messages** (to stderr):
- File not found: `Error: file not found: <path>`
- Wrong type: `Error: unsupported file type (expected .pptx): <path>`
- Corrupt: `Error: failed to open .pptx: <details>`
- Write failure: `Error: failed to write output file <path>: <details>`

Use Python's `logging` module (not `print`) for internal logging. Respect `--verbose` flag.

## Output Format

**Console output** (default, human-readable):
```
Analyzing: example_deck.pptx

Total slides: 15

Ranked by media size (descending):

#1  Slide  5  |  8.2 MB  |  title="Overview of Architecture"
#2  Slide 12  |  3.7 MB  |  title="Customer Deployment"
#3  Slide  2  |  1.1 MB  |  title="Agenda"
...
```

- Use 1-based slide indices
- Format sizes sensibly (KB/MB)
- Show slide title or "(no title)"
- Respect --top N to limit output

**JSON output**: List of SlideMediaStats objects, sorted by total_media_bytes descending

**CSV output**: Flat table with columns: rank, slide_index, slide_title, total_media_bytes, image_bytes, video_bytes, audio_bytes, other_media_bytes

## Code Quality Standards

- Type hints on public functions
- Docstrings for main functions
- Modular structure (analysis logic reusable as library)
- No hard-coded paths
- Support large files (500+ slides, hundreds of MB) efficiently
- Single-pass over slides
