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
