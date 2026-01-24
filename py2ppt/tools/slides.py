"""Slide manipulation tool functions.

Functions for adding, deleting, and reordering slides.
"""

from __future__ import annotations

from ..core.presentation import Presentation


def add_slide(
    presentation: Presentation,
    layout: str | int = "Title and Content",
    *,
    position: int | None = None,
) -> int:
    """Add a new slide to the presentation.

    Args:
        presentation: The presentation to modify
        layout: Layout name (fuzzy matched) or index (0-indexed).
                Common layouts: "Title Slide", "Title and Content",
                "Section Header", "Two Content", "Blank", "Title Only"
        position: Insert position (1-indexed). None = append at end.

    Returns:
        The slide number of the new slide (1-indexed)

    Example:
        >>> slide_num = add_slide(pres, layout="Title Slide")
        >>> slide_num = add_slide(pres, layout="Title and Content", position=2)
    """
    slide = presentation.add_slide(layout=layout, position=position)
    return slide.number


def delete_slide(
    presentation: Presentation,
    slide_number: int,
) -> bool:
    """Delete a slide from the presentation.

    Args:
        presentation: The presentation to modify
        slide_number: The slide number to delete (1-indexed)

    Returns:
        True if deleted successfully, False if slide not found

    Example:
        >>> delete_slide(pres, slide_number=3)
        True
    """
    return presentation.delete_slide(slide_number)


def duplicate_slide(
    presentation: Presentation,
    slide_number: int,
) -> int:
    """Duplicate a slide.

    The duplicated slide is inserted immediately after the original.

    Args:
        presentation: The presentation to modify
        slide_number: The slide number to duplicate (1-indexed)

    Returns:
        The slide number of the new (duplicated) slide

    Raises:
        SlideNotFoundError: If slide number is out of range

    Example:
        >>> new_num = duplicate_slide(pres, slide_number=2)
    """
    new_slide = presentation.duplicate_slide(slide_number)
    return new_slide.number


def reorder_slides(
    presentation: Presentation,
    order: list[int],
) -> None:
    """Reorder slides in the presentation.

    Args:
        presentation: The presentation to modify
        order: New order as list of slide numbers (1-indexed).
               Must contain all slide numbers exactly once.
               e.g., [2, 1, 3] moves slide 2 to first position

    Raises:
        ValueError: If order is invalid

    Example:
        >>> reorder_slides(pres, order=[3, 1, 2])  # Move slide 3 to first
    """
    presentation.reorder_slides(order)
