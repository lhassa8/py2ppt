"""Inspection tool functions.

Functions for examining presentation structure and content.
These are typically called first by AI agents to understand
the template/presentation before making modifications.
"""

from __future__ import annotations

from typing import Any

from ..core.presentation import Presentation


def list_layouts(presentation: Presentation) -> list[dict[str, Any]]:
    """List all available layouts in the presentation.

    Args:
        presentation: The presentation to inspect

    Returns:
        List of layout information dicts:
        [
            {
                "name": "Title Slide",
                "index": 0,
                "placeholders": ["title", "subtitle"]
            },
            {
                "name": "Title and Content",
                "index": 1,
                "placeholders": ["title", "body"]
            },
            ...
        ]

    Example:
        >>> layouts = list_layouts(pres)
        >>> for layout in layouts:
        ...     print(f"{layout['name']}: {layout['placeholders']}")
    """
    layouts = presentation.get_layouts()

    result = []
    for layout in layouts:
        # Convert placeholder objects to type names
        ph_names = []
        for ph in layout.placeholders:
            name = ph.type
            if ph.idx is not None and ph.idx > 0:
                name = f"{name}_{ph.idx}"
            ph_names.append(name)

        result.append({
            "name": layout.name,
            "index": layout.index,
            "placeholders": ph_names,
        })

    return result


def describe_slide(
    presentation: Presentation,
    slide_number: int,
) -> dict[str, Any]:
    """Get detailed information about a slide.

    Args:
        presentation: The presentation to inspect
        slide_number: The slide number (1-indexed)

    Returns:
        Dict with slide information:
        {
            "slide_number": 2,
            "placeholders": {
                "title": "Current Title",
                "body": ["Bullet 1", "Bullet 2"]
            },
            "shapes": [
                {"type": "image", "name": "Picture 1"},
                {"type": "table", "name": "Table 1", "rows": 3, "cols": 4}
            ]
        }

    Example:
        >>> info = describe_slide(pres, 2)
        >>> print(f"Title: {info['placeholders'].get('title')}")
    """
    slide = presentation.get_slide(slide_number)
    return slide.describe()


def get_placeholders(
    presentation: Presentation,
    slide_number: int,
) -> dict[str, str]:
    """Get all placeholder content from a slide.

    Args:
        presentation: The presentation to inspect
        slide_number: The slide number (1-indexed)

    Returns:
        Dict mapping placeholder type to current content:
        {
            "title": "Slide Title",
            "body": "Bullet 1\\nBullet 2\\nBullet 3"
        }

    Example:
        >>> placeholders = get_placeholders(pres, 1)
        >>> print(placeholders.get("title"))
    """
    slide = presentation.get_slide(slide_number)

    result = {}
    for ph_type, shape in slide.get_placeholders().items():
        if shape.text_frame:
            result[ph_type] = shape.text_frame.text
        else:
            result[ph_type] = ""

    return result


def get_theme_colors(presentation: Presentation) -> dict[str, str]:
    """Get theme colors from the presentation.

    Args:
        presentation: The presentation to inspect

    Returns:
        Dict mapping color name to hex value:
        {
            "accent1": "#4472C4",
            "accent2": "#ED7D31",
            "accent3": "#A5A5A5",
            "accent4": "#FFC000",
            "accent5": "#5B9BD5",
            "accent6": "#70AD47",
            "dk1": "#000000",
            "lt1": "#FFFFFF",
            "dk2": "#44546A",
            "lt2": "#E7E6E6",
            "hlink": "#0563C1",
            "folHlink": "#954F72"
        }

    Example:
        >>> colors = get_theme_colors(pres)
        >>> primary = colors.get("accent1")
    """
    return presentation.get_theme_colors()


def get_theme_fonts(presentation: Presentation) -> dict[str, str]:
    """Get theme fonts from the presentation.

    Args:
        presentation: The presentation to inspect

    Returns:
        Dict with heading and body fonts:
        {
            "heading": "Calibri Light",
            "body": "Calibri"
        }

    Example:
        >>> fonts = get_theme_fonts(pres)
        >>> heading_font = fonts["heading"]
    """
    return presentation.get_theme_fonts()


def get_slide_count(presentation: Presentation) -> int:
    """Get the number of slides in the presentation.

    Args:
        presentation: The presentation to inspect

    Returns:
        Number of slides

    Example:
        >>> count = get_slide_count(pres)
        >>> print(f"Presentation has {count} slides")
    """
    return presentation.slide_count


# ============================================================================
# Thumbnails
# ============================================================================


def get_presentation_thumbnail(presentation: Presentation) -> bytes | None:
    """Get the embedded thumbnail image from the presentation.

    PowerPoint files often contain an embedded thumbnail image.
    This function extracts it if available.

    Args:
        presentation: The presentation to inspect

    Returns:
        Image data as bytes, or None if no thumbnail is embedded

    Example:
        >>> thumbnail = get_presentation_thumbnail(pres)
        >>> if thumbnail:
        ...     with open("thumb.png", "wb") as f:
        ...         f.write(thumbnail)
    """
    pkg = presentation._package

    # Try common thumbnail locations in PPTX
    thumbnail_paths = [
        "docProps/thumbnail.jpeg",
        "docProps/thumbnail.png",
        "_rels/.rels",  # Check for thumbnail relationship
    ]

    for path in thumbnail_paths[:2]:  # Check direct paths first
        data = pkg.get_part(path)
        if data:
            return data

    return None


def get_slide_thumbnail(
    presentation: Presentation,
    slide_number: int,
    *,
    width: int = 320,
    height: int = 240,
    format: str = "png",
) -> bytes | None:
    """Generate a thumbnail image for a specific slide.

    Requires LibreOffice and pdftoppm/ImageMagick for rendering.
    For quick previews without external dependencies, use
    get_presentation_thumbnail() instead.

    Args:
        presentation: The presentation to process
        slide_number: The slide number (1-indexed)
        width: Thumbnail width in pixels (default 320)
        height: Thumbnail height in pixels (default 240)
        format: Image format - "png" or "jpg" (default "png")

    Returns:
        Image data as bytes, or None if generation failed

    Example:
        >>> thumb = get_slide_thumbnail(pres, 1, width=200, height=150)
        >>> if thumb:
        ...     with open("slide1_thumb.png", "wb") as f:
        ...         f.write(thumb)
    """
    import tempfile
    from pathlib import Path

    from .export import export_slide_to_image

    # Validate slide number
    if slide_number < 1 or slide_number > presentation.slide_count:
        return None

    # Create temporary file for the thumbnail
    with tempfile.NamedTemporaryFile(
        suffix=f".{format}", delete=False
    ) as tmp_file:
        tmp_path = Path(tmp_file.name)

    try:
        # Use export function to generate thumbnail
        export_slide_to_image(
            presentation,
            slide_number,
            tmp_path,
            format=format,
            width=width,
            height=height,
        )

        # Read the generated image
        if tmp_path.exists():
            with open(tmp_path, "rb") as f:
                return f.read()
        return None

    except RuntimeError:
        # Export tools not available
        return None

    finally:
        # Clean up temp file
        if tmp_path.exists():
            tmp_path.unlink()


def save_slide_thumbnail(
    presentation: Presentation,
    slide_number: int,
    output_path: str,
    *,
    width: int = 320,
    height: int = 240,
) -> bool:
    """Save a slide thumbnail to a file.

    Requires LibreOffice and pdftoppm/ImageMagick for rendering.

    Args:
        presentation: The presentation to process
        slide_number: The slide number (1-indexed)
        output_path: Path for the output image file
        width: Thumbnail width in pixels (default 320)
        height: Thumbnail height in pixels (default 240)

    Returns:
        True if saved successfully, False otherwise

    Example:
        >>> success = save_slide_thumbnail(pres, 1, "thumb.png")
        >>> if success:
        ...     print("Thumbnail saved!")
    """
    from pathlib import Path

    from .export import export_slide_to_image

    # Validate slide number
    if slide_number < 1 or slide_number > presentation.slide_count:
        return False

    output = Path(output_path)

    # Determine format from extension
    ext = output.suffix.lower().lstrip(".")
    if ext in ("jpg", "jpeg"):
        fmt = "jpg"
    elif ext == "gif":
        fmt = "gif"
    else:
        fmt = "png"

    try:
        export_slide_to_image(
            presentation,
            slide_number,
            output,
            format=fmt,
            width=width,
            height=height,
        )
        return output.exists()

    except RuntimeError:
        return False


def get_all_thumbnails(
    presentation: Presentation,
    output_dir: str | None = None,
    *,
    width: int = 320,
    height: int = 240,
    format: str = "png",
) -> list[bytes] | list[str]:
    """Generate thumbnails for all slides.

    If output_dir is provided, saves files and returns paths.
    Otherwise, returns list of image data as bytes.

    Requires LibreOffice and pdftoppm/ImageMagick for rendering.

    Args:
        presentation: The presentation to process
        output_dir: Directory to save thumbnails (optional)
        width: Thumbnail width in pixels (default 320)
        height: Thumbnail height in pixels (default 240)
        format: Image format - "png" or "jpg" (default "png")

    Returns:
        If output_dir is None: List of image data as bytes
        If output_dir is provided: List of output file paths

    Example:
        >>> # Get as bytes
        >>> thumbnails = get_all_thumbnails(pres, width=200, height=150)
        >>> for i, thumb in enumerate(thumbnails):
        ...     print(f"Slide {i+1}: {len(thumb)} bytes")

        >>> # Save to files
        >>> paths = get_all_thumbnails(pres, "thumbs/", format="jpg")
        >>> for path in paths:
        ...     print(f"Saved: {path}")
    """
    from pathlib import Path

    from .export import export_all_slides

    num_slides = presentation.slide_count
    if num_slides == 0:
        return []

    if output_dir is not None:
        # Save to files
        output = Path(output_dir)
        output.mkdir(parents=True, exist_ok=True)

        try:
            paths = export_all_slides(
                presentation,
                output,
                format=format,
                prefix="thumb",
                width=width,
                height=height,
            )
            return [str(p) for p in paths]

        except RuntimeError:
            return []

    else:
        # Return as bytes
        results = []
        for slide_num in range(1, num_slides + 1):
            thumb = get_slide_thumbnail(
                presentation,
                slide_num,
                width=width,
                height=height,
                format=format,
            )
            if thumb:
                results.append(thumb)

        return results
