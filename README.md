# slide_size

⸻

Project: PowerPoint “Heavy Slides” Analyzer

1. Purpose and goals

We want a small Python tool that:
	1.	Analyzes a .pptx file.
	2.	Estimates how “heavy” each slide is, primarily due to:
	•	Raster images (PNG, JPEG, etc.).
	•	Embedded media (video, audio).
	3.	Produces a ranked list of slides by their estimated media size.
	4.	Optionally outputs a simple report (console and/or CSV/JSON).

The primary use case is to quickly identify slides that are contributing most to the PPTX file size, so they can be optimized.

This is not about pixel dimensions or visual complexity—it’s about actual bytes of images/media embedded in the file.

⸻

2. Scope

In scope
	•	.pptx files (Office Open XML).
	•	On-slide content:
	•	Pictures (bitmap images).
	•	Embedded media (video, audio).
	•	A CLI (command-line interface) for:
	•	Taking a .pptx file as input.
	•	Producing a ranked summary (largest slides first).
	•	Basic error handling and sensible exit codes.
	•	A small test suite with a few synthetic PPTX files.

Out of scope (for now)
	•	Legacy .ppt binaries.
	•	SmartArt, vector drawings, shapes without images.
	•	Embedded documents (Excel, Word) – count them if easy via the same mechanism, but not required.
	•	Integration with PowerPoint itself (no COM automation, no Office JS).

⸻

3. Environment & dependencies
	•	Language: Python 3.10+ (assume 3.10 or later).
	•	Target OS: Should run on at least:
	•	Linux
	•	macOS
	•	Windows
	•	Allowed dependencies:
	•	python-pptx (preferred)
	•	Standard library modules (zipfile, argparse, logging, json, csv, etc.)

If you can do the core functionality using only the standard library (zipfile + XML parsing), that’s a bonus, but not required for the first version.

⸻

4. Functional requirements

4.1 Command-line interface

Implement a CLI entry point, e.g.:

python pptx_heavy_slides.py <path-to-pptx> [options]

Required arguments
	•	input_path (positional):
	•	Path to a .pptx file.
	•	If file does not exist or is not a valid .pptx, print a clear error to stderr and exit with non-zero status.

Optional arguments
	•	--top N
	•	Integer; show only the top N heaviest slides.
	•	Default: show all slides.
	•	--output-json <path>
	•	Write full results as JSON to the given path.
	•	--output-csv <path>
	•	Write full results as CSV to the given path.
	•	--include-shared-media / --ignore-shared-media
	•	Behavior:
	•	--ignore-shared-media (default):
If an image/media file is used on multiple slides, count its bytes only once, attributed to the first slide on which it appears. Subsequent slides using the same media should show 0 bytes for that specific media item (but still include other unique media, if any).
	•	--include-shared-media:
Count the media’s bytes on every slide that uses it. This answers “how heavy is this slide if you think of it alone?” but doesn’t match storage contribution.
	•	--verbose
	•	Increase logging verbosity (e.g., debug messages about what’s being parsed).
	•	--version
	•	Output tool version and exit.

⸻

4.2 Output format (console)

Default console output should be human-readable, something like:

Analyzing: example_deck.pptx

Total slides: 15

Ranked by media size (descending):

#1  Slide  5  |  8.2 MB  |  title="Overview of Architecture"
#2  Slide 12  |  3.7 MB  |  title="Customer Deployment"
#3  Slide  2  |  1.1 MB  |  title="Agenda"
...
#15 Slide  9  |  0.0 MB  |  title="Summary"

Details:
	•	Use 1-based slide index (as users see in PowerPoint).
	•	Show:
	•	Rank.
	•	Slide number.
	•	Total media bytes (converted to KB/MB with sensible formatting).
	•	Slide title, if available; otherwise something like "(no title)".

If --top N is used, display only that many rows.

⸻

4.3 Programmatic data model

Internally, define a simple structure for a slide’s analysis:

class SlideMediaStats(TypedDict):
    slide_index: int            # 1-based index
    slide_title: str | None
    total_media_bytes: int
    image_bytes: int            # subset of total_media_bytes
    video_bytes: int            # subset of total_media_bytes
    audio_bytes: int            # subset of total_media_bytes
    other_media_bytes: int      # optional, for embedded stuff not image/video/audio
    media_items: list[dict]     # see below

Each media_items element:

{
    "type": "image" | "video" | "audio" | "other",
    "size_bytes": int,
    "filename": str | None,        # e.g., "image3.png" from /ppt/media
    "content_type": str | None,    # e.g., "image/png", "video/mp4"
    "relationship_id": str | None, # rId from relationships, if easily available
    "shared": bool                 # True if used on multiple slides
}

JSON output (--output-json) should basically be a list of these objects, sorted by total_media_bytes descending.

CSV output (--output-csv) can be a flat table with columns:
	•	rank
	•	slide_index
	•	slide_title
	•	total_media_bytes
	•	image_bytes
	•	video_bytes
	•	audio_bytes
	•	other_media_bytes

(We don’t need to serialize media_items into CSV for this exercise.)

⸻

5. Algorithm / implementation approach

5.1 High-level steps
	1.	Open the PPTX
	•	Verify the file extension is .pptx.
	•	Try to load it with python-pptx’s Presentation (or, alternatively, with zipfile).
	•	On failure, print error and exit.
	2.	Enumerate slides
	•	For each slide:
	•	Determine slide index (1-based).
	•	Extract a sensible “title”:
	•	Try to use the title placeholder if available (e.g., shape.is_placeholder and shape.placeholder_format.type).
	•	If no title, fallback to None or "(no title)" in display.
	3.	Identify media items
You have two possible strategies; either is acceptable as long as it’s implemented correctly and clearly:
Strategy A – via python-pptx shape objects
For each slide:
	•	Iterate over slide.shapes.
	•	If shape.shape_type is PICTURE, treat as an image:
	•	Extract bytes via shape.image.blob.
	•	Content type can often be shape.image.content_type.
	•	For media (audio/video), inspect:
	•	shape.media_format or related attributes/relationships.
	•	If needed, fall back to looking at slide’s related parts in the underlying package (more advanced).
	•	Build a list of media items for the slide.
Strategy B – via ZIP / relationships (more exact, advanced)
	•	Use zipfile to inspect /ppt/slides/slideN.xml.
	•	Parse XML to find <a:blip> references which point via relationship IDs (r:embed) to media in /ppt/media/*.
	•	Use /ppt/slides/_rels/slideN.xml.rels to map rIdX → /ppt/media/imageY.ext or video/audio.
	•	Use ZipFile.getinfo() to get the exact file size of each /ppt/media/* entry.
For this exercise, Strategy A with python-pptx is fine, but keep the design open for Strategy B as a later enhancement.
	4.	Shared media detection
We need to distinguish media files that are used on multiple slides.
Approach:
	•	Create a global dictionary keyed by media identifier, e.g.:

media_map: dict[str, dict]  # key could be filename or partname


	•	For each media item (image/audio/video):
	•	Determine its “partname” or “filename” inside the package, e.g. /ppt/media/image3.png.
	•	Use that as the key.
	•	Track:
	•	Total size in bytes (from blob or zip entry).
	•	Slides on which it appears.
	•	After processing all slides:
	•	Mark each media item as shared (True) if it appears on > 1 slide.
	•	When building SlideMediaStats, apply the logic:
	•	If --ignore-shared-media:
	•	Only count bytes for that media on the first slide it appears on.
	•	If --include-shared-media:
	•	Count bytes on every slide where it appears.

	5.	Calculate per-slide statistics
For each slide, based on its list of media items after applying shared-media rules:
	•	image_bytes = sum sizes where type == "image".
	•	video_bytes = sum sizes where type == "video".
	•	audio_bytes = sum sizes where type == "audio".
	•	other_media_bytes = sum sizes where type == "other".
	•	total_media_bytes = sum of all of the above.
	6.	Sort and rank
	•	Sort the list of SlideMediaStats by total_media_bytes descending.
	•	Assign rank = enumeration index after sorting (starting at 1).
	7.	Output
	•	Print summary to stdout, respecting --top.
	•	If --output-json is provided:
	•	Write JSON (UTF-8, pretty-printed).
	•	If --output-csv is provided:
	•	Write CSV with header row.

⸻

6. Error handling and logging

6.1 Errors
	•	File not found:
	•	Print a clear error: Error: file not found: <path>.
	•	Exit with status code 1.
	•	Not a .pptx file:
	•	Print: Error: unsupported file type (expected .pptx): <path>.
	•	Exit status 1.
	•	Corrupt or invalid PPTX:
	•	Print: Error: failed to open .pptx: <details>.
	•	Exit status 1.
	•	Output file write failures (JSON/CSV):
	•	Print: Error: failed to write output file <path>: <details>.
	•	Exit status 1.

6.2 Logging
	•	By default, minimal logging: only high-level messages.
	•	If --verbose is set:
	•	Log extra info such as:
	•	“Found N slides.”
	•	“Slide 3: found 2 images, 1 video.”
	•	“Media /ppt/media/image3.png used on slides [2, 7].”

Use Python’s logging module (not print) for internal logging.

⸻

7. Testing requirements

Create a minimal automated test suite (e.g., using pytest or unittest) with at least:
	1.	Single-slide deck with one image
	•	Expect: slide 1 has non-zero image bytes; total_media_bytes matches that value.
	2.	Multi-slide deck, same image on multiple slides
	•	With --ignore-shared-media logic:
	•	Only first slide where image appears has non-zero size for that media.
	•	With --include-shared-media logic:
	•	All slides using the image show that size.
	3.	Deck with one video and one audio
	•	Ensure they are categorized correctly and included in total.
	4.	Empty deck (no images/media)
	•	All slides should have total_media_bytes = 0.
	5.	Invalid input
	•	Non-existent file, wrong extension, or corrupt PPTX should produce errors and non-zero exit code.

You can generate small PPTX files using PowerPoint or python-pptx in a test setup script.

⸻

8. Non-functional requirements
	•	Performance:
	•	Should comfortably handle decks up to ~500 slides and multi-hundred-MB PPTXs without excessive memory usage.
	•	Single-pass over slides is fine.
	•	Code quality:
	•	Clear, modular structure.
	•	Type hints where helpful.
	•	Docstrings for the main functions.
	•	No hard-coded paths; everything controlled via CLI parameters.
	•	Maintainability:
	•	Keep slide-analysis logic separate from CLI parsing so it could be imported and used as a library function in the future.
	•	e.g., a function like:

def analyze_pptx_media(path: str, include_shared_media: bool = False) -> list[SlideMediaStats]:
    ...



⸻

9. Deliverables
	1.	Python module/script, e.g. pptx_heavy_slides.py, containing:
	•	analyze_pptx_media(...) core function.
	•	CLI entry point (under if __name__ == "__main__":).
	2.	Test suite:
	•	In a tests/ directory (or similar).
	•	Instructions or script to run the tests (pytest or python -m unittest).
	3.	Short README.md describing:
	•	What the tool does.
	•	How to install dependencies.
	•	Example commands and sample output.

⸻

