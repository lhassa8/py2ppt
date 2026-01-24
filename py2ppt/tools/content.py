"""Content manipulation tool functions.

Functions for setting text content in slides.
"""

from __future__ import annotations

from typing import List, Optional, Union

from ..core.presentation import Presentation
from ..core.slide import Slide
from ..oxml.shapes import Shape


def set_title(
    presentation: Presentation,
    slide_number: int,
    text: str,
    *,
    font_size: Optional[int] = None,
    font_family: Optional[str] = None,
    bold: bool = False,
    color: Optional[str] = None,
) -> None:
    """Set the title of a slide.

    Args:
        presentation: The presentation to modify
        slide_number: The slide number (1-indexed)
        text: The title text
        font_size: Font size in points (e.g., 32)
        font_family: Font family name (e.g., "Arial")
        bold: Whether to make the text bold
        color: Color as hex ("#FF0000"), rgb("rgb(255,0,0)"), or name ("red")

    Example:
        >>> set_title(pres, 1, "Q4 Business Review")
        >>> set_title(pres, 2, "Key Metrics", color="#0066CC", bold=True)
    """
    slide = presentation.get_slide(slide_number)
    slide.set_title(
        text,
        font_size=font_size,
        font_family=font_family,
        bold=bold,
        color=color,
    )


def set_subtitle(
    presentation: Presentation,
    slide_number: int,
    text: str,
    *,
    font_size: Optional[int] = None,
    font_family: Optional[str] = None,
    bold: bool = False,
    color: Optional[str] = None,
) -> None:
    """Set the subtitle of a slide.

    Args:
        presentation: The presentation to modify
        slide_number: The slide number (1-indexed)
        text: The subtitle text
        font_size: Font size in points
        font_family: Font family name
        bold: Whether to make the text bold
        color: Color as hex, rgb, or name

    Example:
        >>> set_subtitle(pres, 1, "Prepared by Analytics Team")
    """
    slide = presentation.get_slide(slide_number)
    slide.set_subtitle(
        text,
        font_size=font_size,
        font_family=font_family,
        bold=bold,
        color=color,
    )


def set_body(
    presentation: Presentation,
    slide_number: int,
    content: Union[str, List[str]],
    *,
    levels: Optional[List[int]] = None,
    font_size: Optional[int] = None,
    font_family: Optional[str] = None,
    color: Optional[str] = None,
) -> None:
    """Set the body content of a slide.

    Args:
        presentation: The presentation to modify
        slide_number: The slide number (1-indexed)
        content: Single string or list of bullet points
        levels: Optional list of indent levels (0-8) for each bullet.
                Default is 0 (top level) for all items.
        font_size: Font size in points
        font_family: Font family name
        color: Color as hex, rgb, or name

    Example:
        >>> set_body(pres, 2, [
        ...     "Revenue up 20%",
        ...     "New markets opened",
        ...     "Customer satisfaction at 95%"
        ... ])
        >>> # With nested bullets:
        >>> set_body(pres, 3, [
        ...     "Main point",
        ...     "Sub-point 1",
        ...     "Sub-point 2",
        ...     "Another main point"
        ... ], levels=[0, 1, 1, 0])
    """
    slide = presentation.get_slide(slide_number)
    slide.set_body(
        content,
        levels=levels,
        font_size=font_size,
        font_family=font_family,
        color=color,
    )


def add_bullet(
    presentation: Presentation,
    slide_number: int,
    text: str,
    *,
    level: int = 0,
    font_size: Optional[int] = None,
    font_family: Optional[str] = None,
    color: Optional[str] = None,
) -> None:
    """Add a bullet point to the slide body.

    Appends a new bullet to existing body content.

    Args:
        presentation: The presentation to modify
        slide_number: The slide number (1-indexed)
        text: The bullet text
        level: Indent level (0-8). 0 is top level.
        font_size: Font size in points
        font_family: Font family name
        color: Color as hex, rgb, or name

    Example:
        >>> add_bullet(pres, 2, "Additional point")
        >>> add_bullet(pres, 2, "Sub-point", level=1)
    """
    slide = presentation.get_slide(slide_number)
    slide.add_bullet(
        text,
        level=level,
        font_size=font_size,
        font_family=font_family,
        color=color,
    )


def set_placeholder_text(
    presentation: Presentation,
    slide_number: int,
    placeholder: str,
    text: str,
    *,
    font_size: Optional[int] = None,
    font_family: Optional[str] = None,
    bold: bool = False,
    color: Optional[str] = None,
) -> None:
    """Set text in a specific placeholder.

    Args:
        presentation: The presentation to modify
        slide_number: The slide number (1-indexed)
        placeholder: Placeholder type or name. Common values:
                    "title", "subtitle", "body", "content",
                    "footer", "date", "slide_number"
                    For multiple placeholders of same type: "body_1", "body_2"
        text: The text content
        font_size: Font size in points
        font_family: Font family name
        bold: Whether to make the text bold
        color: Color as hex, rgb, or name

    Example:
        >>> set_placeholder_text(pres, 2, "body_1", "Left column content")
        >>> set_placeholder_text(pres, 2, "body_2", "Right column content")
    """
    slide = presentation.get_slide(slide_number)
    slide.set_placeholder_text(
        placeholder,
        text,
        font_size=font_size,
        font_family=font_family,
        bold=bold,
        color=color,
    )


def add_text_box(
    presentation: Presentation,
    slide_number: int,
    text: str,
    left: Union[str, int],
    top: Union[str, int],
    width: Union[str, int],
    height: Union[str, int],
    *,
    font_size: Optional[int] = None,
    font_family: Optional[str] = None,
    bold: bool = False,
    color: Optional[str] = None,
) -> None:
    """Add a text box at a specific position.

    Args:
        presentation: The presentation to modify
        slide_number: The slide number (1-indexed)
        text: The text content
        left: Left position (e.g., "1in", "2.5cm", or EMU value)
        top: Top position
        width: Width
        height: Height
        font_size: Font size in points
        font_family: Font family name
        bold: Whether to make the text bold
        color: Color as hex, rgb, or name

    Example:
        >>> add_text_box(pres, 1, "Note", "1in", "6in", "2in", "0.5in")
    """
    slide = presentation.get_slide(slide_number)
    slide.add_text_box(
        text,
        left,
        top,
        width,
        height,
        font_size=font_size,
        font_family=font_family,
        bold=bold,
        color=color,
    )
