#!/usr/bin/env python3
"""
PowerPoint Heavy Slides Analyzer

Analyzes .pptx files to identify which slides contribute most to file size
due to embedded media (images, videos, audio).
"""

import argparse
import csv
import json
import logging
import sys
from pathlib import Path
from typing import TypedDict
from collections import defaultdict

from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE


__version__ = "1.0.0"


# Data model
class MediaItem(TypedDict):
    """Represents a single media item on a slide."""
    type: str  # "image", "video", "audio", "other"
    size_bytes: int
    filename: str | None
    content_type: str | None
    relationship_id: str | None
    shared: bool


class SlideMediaStats(TypedDict):
    """Statistics about media content on a single slide."""
    slide_index: int  # 1-based index
    slide_title: str | None
    total_media_bytes: int
    image_bytes: int
    video_bytes: int
    audio_bytes: int
    other_media_bytes: int
    media_items: list[MediaItem]


def setup_logging(verbose: bool = False) -> None:
    """Configure logging based on verbosity level."""
    level = logging.DEBUG if verbose else logging.INFO
    logging.basicConfig(
        level=level,
        format='%(levelname)s: %(message)s'
    )


def get_slide_title(slide) -> str | None:
    """
    Extract the title from a slide.

    Args:
        slide: A slide object from python-pptx

    Returns:
        The slide title text, or None if no title found
    """
    if not slide.shapes.title:
        return None

    try:
        title_text = slide.shapes.title.text.strip()
        return title_text if title_text else None
    except (AttributeError, IndexError):
        return None


def analyze_pptx_media(path: str, include_shared_media: bool = False) -> list[SlideMediaStats]:
    """
    Analyze a PowerPoint file to determine media size per slide.

    Args:
        path: Path to the .pptx file
        include_shared_media: If True, count shared media on every slide.
                             If False (default), count shared media only on first appearance.

    Returns:
        List of SlideMediaStats, one per slide, sorted by total_media_bytes descending

    Raises:
        FileNotFoundError: If the file doesn't exist
        ValueError: If the file is not a .pptx file or is corrupt
    """
    # Validate file
    file_path = Path(path)
    if not file_path.exists():
        raise FileNotFoundError(f"file not found: {path}")

    if file_path.suffix.lower() != '.pptx':
        raise ValueError(f"unsupported file type (expected .pptx): {path}")

    # Try to open the presentation
    try:
        prs = Presentation(path)
    except Exception as e:
        raise ValueError(f"failed to open .pptx: {e}")

    logging.info(f"Found {len(prs.slides)} slides")

    # Track all media across slides to detect sharing
    media_registry = {}  # key: image blob hash or media part name -> value: {size, slides, first_slide}
    slide_media_map = defaultdict(list)  # slide_index -> list of media items

    # First pass: collect all media and track where it appears
    for slide_idx, slide in enumerate(prs.slides, start=1):
        logging.debug(f"Analyzing slide {slide_idx}")

        for shape in slide.shapes:
            # Handle images
            if shape.shape_type == MSO_SHAPE_TYPE.PICTURE:
                try:
                    image = shape.image
                    blob = image.blob
                    size_bytes = len(blob)
                    content_type = image.content_type

                    # Use blob hash as identifier for detecting duplicates
                    media_key = hash(blob)

                    if media_key not in media_registry:
                        media_registry[media_key] = {
                            'size': size_bytes,
                            'type': 'image',
                            'content_type': content_type,
                            'filename': image.filename if hasattr(image, 'filename') else None,
                            'slides': [],
                            'first_slide': slide_idx
                        }

                    media_registry[media_key]['slides'].append(slide_idx)

                    # Store reference for this slide
                    slide_media_map[slide_idx].append({
                        'media_key': media_key,
                        'type': 'image',
                        'size_bytes': size_bytes,
                        'content_type': content_type,
                        'filename': media_registry[media_key]['filename']
                    })

                    logging.debug(f"  Found image: {size_bytes} bytes, type={content_type}")

                except Exception as e:
                    logging.warning(f"  Failed to extract image from slide {slide_idx}: {e}")

            # Handle video and audio
            elif hasattr(shape, 'shape_type') and shape.shape_type == MSO_SHAPE_TYPE.MEDIA:
                try:
                    media_format = shape.media_format
                    if media_format:
                        # Get the media part
                        media_part = shape.part.related_part(shape._element.xpath('.//a:videoFile | .//a:audioFile | .//p:media')[0].get('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}embed'))

                        size_bytes = len(media_part.blob)
                        content_type = media_part.content_type

                        # Determine media type
                        if 'video' in content_type.lower():
                            media_type = 'video'
                        elif 'audio' in content_type.lower():
                            media_type = 'audio'
                        else:
                            media_type = 'other'

                        # Use part name as identifier
                        media_key = media_part.partname

                        if media_key not in media_registry:
                            media_registry[media_key] = {
                                'size': size_bytes,
                                'type': media_type,
                                'content_type': content_type,
                                'filename': Path(media_part.partname).name,
                                'slides': [],
                                'first_slide': slide_idx
                            }

                        media_registry[media_key]['slides'].append(slide_idx)

                        slide_media_map[slide_idx].append({
                            'media_key': media_key,
                            'type': media_type,
                            'size_bytes': size_bytes,
                            'content_type': content_type,
                            'filename': media_registry[media_key]['filename']
                        })

                        logging.debug(f"  Found {media_type}: {size_bytes} bytes, type={content_type}")

                except Exception as e:
                    logging.warning(f"  Failed to extract media from slide {slide_idx}: {e}")

    # Mark shared media
    for media_key, info in media_registry.items():
        if len(info['slides']) > 1:
            logging.debug(f"Media {info.get('filename', media_key)} appears on slides: {info['slides']}")

    # Second pass: build SlideMediaStats for each slide
    results = []
    for slide_idx, slide in enumerate(prs.slides, start=1):
        title = get_slide_title(slide)
        media_items = []

        image_bytes = 0
        video_bytes = 0
        audio_bytes = 0
        other_media_bytes = 0

        # Process media for this slide
        for media_ref in slide_media_map.get(slide_idx, []):
            media_key = media_ref['media_key']
            media_info = media_registry[media_key]

            is_shared = len(media_info['slides']) > 1
            is_first_appearance = media_info['first_slide'] == slide_idx

            # Determine if we should count the bytes
            if include_shared_media or not is_shared or is_first_appearance:
                count_bytes = media_ref['size_bytes']
            else:
                count_bytes = 0  # Shared media, not first appearance, don't count

            # Create media item
            media_item: MediaItem = {
                'type': media_ref['type'],
                'size_bytes': count_bytes,
                'filename': media_ref['filename'],
                'content_type': media_ref['content_type'],
                'relationship_id': None,  # Could be enhanced later
                'shared': is_shared
            }
            media_items.append(media_item)

            # Accumulate by type
            if media_ref['type'] == 'image':
                image_bytes += count_bytes
            elif media_ref['type'] == 'video':
                video_bytes += count_bytes
            elif media_ref['type'] == 'audio':
                audio_bytes += count_bytes
            else:
                other_media_bytes += count_bytes

        total_media_bytes = image_bytes + video_bytes + audio_bytes + other_media_bytes

        stats: SlideMediaStats = {
            'slide_index': slide_idx,
            'slide_title': title,
            'total_media_bytes': total_media_bytes,
            'image_bytes': image_bytes,
            'video_bytes': video_bytes,
            'audio_bytes': audio_bytes,
            'other_media_bytes': other_media_bytes,
            'media_items': media_items
        }
        results.append(stats)

    # Sort by total_media_bytes descending
    results.sort(key=lambda x: x['total_media_bytes'], reverse=True)

    return results


def format_bytes(num_bytes: int) -> str:
    """Format bytes into human-readable string (KB, MB, GB)."""
    if num_bytes == 0:
        return "0.0 MB"

    for unit in ['B', 'KB', 'MB', 'GB']:
        if num_bytes < 1024.0 or unit == 'GB':
            if unit == 'B':
                return f"{num_bytes} {unit}"
            return f"{num_bytes:.1f} {unit}"
        num_bytes /= 1024.0

    return f"{num_bytes:.1f} GB"


def print_console_output(results: list[SlideMediaStats], filename: str, top_n: int | None = None) -> None:
    """
    Print human-readable console output.

    Args:
        results: List of SlideMediaStats (should be pre-sorted)
        filename: Name of the analyzed file
        top_n: If specified, show only top N slides
    """
    print(f"\nAnalyzing: {filename}")
    print(f"\nTotal slides: {len(results)}")
    print("\nRanked by media size (descending):\n")

    display_results = results[:top_n] if top_n else results

    for rank, stats in enumerate(display_results, start=1):
        slide_num = stats['slide_index']
        size_str = format_bytes(stats['total_media_bytes'])
        title = stats['slide_title'] if stats['slide_title'] else "(no title)"

        print(f"#{rank:<3} Slide {slide_num:<3} | {size_str:>10} | title=\"{title}\"")

    print()


def write_json_output(results: list[SlideMediaStats], output_path: str) -> None:
    """
    Write results to JSON file.

    Args:
        results: List of SlideMediaStats
        output_path: Path to output file

    Raises:
        IOError: If file write fails
    """
    try:
        with open(output_path, 'w', encoding='utf-8') as f:
            json.dump(results, f, indent=2, ensure_ascii=False)
        logging.info(f"JSON output written to: {output_path}")
    except Exception as e:
        raise IOError(f"failed to write output file {output_path}: {e}")


def write_csv_output(results: list[SlideMediaStats], output_path: str) -> None:
    """
    Write results to CSV file.

    Args:
        results: List of SlideMediaStats
        output_path: Path to output file

    Raises:
        IOError: If file write fails
    """
    try:
        with open(output_path, 'w', newline='', encoding='utf-8') as f:
            writer = csv.writer(f)
            # Header
            writer.writerow([
                'rank',
                'slide_index',
                'slide_title',
                'total_media_bytes',
                'image_bytes',
                'video_bytes',
                'audio_bytes',
                'other_media_bytes'
            ])

            # Data rows
            for rank, stats in enumerate(results, start=1):
                writer.writerow([
                    rank,
                    stats['slide_index'],
                    stats['slide_title'] or '',
                    stats['total_media_bytes'],
                    stats['image_bytes'],
                    stats['video_bytes'],
                    stats['audio_bytes'],
                    stats['other_media_bytes']
                ])

        logging.info(f"CSV output written to: {output_path}")
    except Exception as e:
        raise IOError(f"failed to write output file {output_path}: {e}")


def main() -> int:
    """Main CLI entry point. Returns exit code."""
    parser = argparse.ArgumentParser(
        description='Analyze PowerPoint files to identify heavy slides with embedded media.',
        formatter_class=argparse.RawDescriptionHelpFormatter
    )

    parser.add_argument(
        'input_path',
        help='Path to the .pptx file to analyze'
    )

    parser.add_argument(
        '--top',
        type=int,
        metavar='N',
        help='Show only the top N heaviest slides'
    )

    parser.add_argument(
        '--output-json',
        metavar='PATH',
        help='Write results as JSON to the specified path'
    )

    parser.add_argument(
        '--output-csv',
        metavar='PATH',
        help='Write results as CSV to the specified path'
    )

    media_group = parser.add_mutually_exclusive_group()
    media_group.add_argument(
        '--include-shared-media',
        action='store_true',
        help='Count shared media on every slide that uses it'
    )
    media_group.add_argument(
        '--ignore-shared-media',
        action='store_true',
        default=True,
        help='Count shared media only on first slide (default)'
    )

    parser.add_argument(
        '--verbose',
        action='store_true',
        help='Enable verbose debug logging'
    )

    parser.add_argument(
        '--version',
        action='version',
        version=f'%(prog)s {__version__}'
    )

    args = parser.parse_args()

    # Setup logging
    setup_logging(args.verbose)

    try:
        # Analyze the presentation
        results = analyze_pptx_media(
            args.input_path,
            include_shared_media=args.include_shared_media
        )

        # Console output (always show)
        print_console_output(results, args.input_path, args.top)

        # Optional JSON output
        if args.output_json:
            write_json_output(results, args.output_json)

        # Optional CSV output
        if args.output_csv:
            write_csv_output(results, args.output_csv)

        return 0

    except FileNotFoundError as e:
        print(f"Error: {e}", file=sys.stderr)
        return 1
    except ValueError as e:
        print(f"Error: {e}", file=sys.stderr)
        return 1
    except IOError as e:
        print(f"Error: {e}", file=sys.stderr)
        return 1
    except Exception as e:
        print(f"Error: unexpected error: {e}", file=sys.stderr)
        logging.exception("Unexpected error occurred")
        return 1


if __name__ == "__main__":
    sys.exit(main())
