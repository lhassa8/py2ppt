"""Style tool functions.

Functions for applying styling to presentation content.
"""

from __future__ import annotations

from typing import Optional

from ..core.presentation import Presentation
from ..oxml.slide import update_slide_in_package
from ..oxml.text import RunProperties
from ..utils.colors import parse_color, is_theme_color


def set_text_style(
    presentation: Presentation,
    slide_number: int,
    placeholder: str,
    *,
    font: Optional[str] = None,
    size: Optional[str] = None,
    color: Optional[str] = None,
    bold: Optional[bool] = None,
    italic: Optional[bool] = None,
    underline: Optional[bool] = None,
) -> None:
    """Set text styling for a placeholder.

    Args:
        presentation: The presentation to modify
        slide_number: The slide number (1-indexed)
        placeholder: Placeholder type (e.g., "title", "body")
        font: Font family name (e.g., "Arial", "Times New Roman")
        size: Font size (e.g., "32pt", "24pt")
        color: Color as hex ("#FF0000"), rgb, name, or theme ("accent1")
        bold: Whether text should be bold
        italic: Whether text should be italic
        underline: Whether text should be underlined

    Example:
        >>> set_text_style(pres, 1, "title",
        ...     font="Arial Black", size="44pt", color="#0066CC", bold=True)
    """
    slide = presentation.get_slide(slide_number)
    shape = slide._find_placeholder(placeholder)

    if shape.text_frame is None:
        return  # No text to style

    # Parse size to centipoints
    font_size = None
    if size:
        size = size.strip().lower()
        if size.endswith("pt"):
            font_size = int(float(size[:-2]) * 100)
        else:
            font_size = int(float(size) * 100)

    # Parse color
    color_val = None
    theme_color = None
    if color:
        parsed = parse_color(color)
        if is_theme_color(parsed):
            theme_color = parsed.split(":")[1]
        else:
            color_val = parsed

    # Apply styling to all runs in all paragraphs
    for para in shape.text_frame.paragraphs:
        for run in para.runs:
            if font:
                run.properties.font_family = font
            if font_size:
                run.properties.font_size = font_size
            if color_val:
                run.properties.color = color_val
            if theme_color:
                run.properties.theme_color = theme_color
            if bold is not None:
                run.properties.bold = bold
            if italic is not None:
                run.properties.italic = italic
            if underline is not None:
                run.properties.underline = underline

    # Save changes
    update_slide_in_package(
        presentation._package,
        slide_number,
        slide._part,
    )
