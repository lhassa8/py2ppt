"""Presentation tool functions.

Functions for creating, opening, and saving presentations.
"""

from __future__ import annotations

from pathlib import Path
from typing import BinaryIO

from ..core.presentation import Presentation


def create_presentation(
    template: str | Path | None = None,
) -> Presentation:
    """Create a new presentation.

    Args:
        template: Optional path to template file. If None, creates blank presentation.

    Returns:
        New Presentation object

    Example:
        >>> pres = create_presentation()  # Blank presentation
        >>> pres = create_presentation(template="corporate.pptx")  # From template
    """
    if template is None:
        return Presentation.new()
    else:
        return Presentation.from_template(template)


def open_presentation(path: str | Path | BinaryIO) -> Presentation:
    """Open an existing presentation.

    Args:
        path: Path to .pptx file or file-like object

    Returns:
        Presentation object

    Example:
        >>> pres = open_presentation("existing.pptx")
    """
    return Presentation.open(path)


def save_presentation(
    presentation: Presentation,
    path: str | Path | BinaryIO,
) -> None:
    """Save a presentation to file.

    Args:
        presentation: The presentation to save
        path: Destination path or file-like object

    Example:
        >>> save_presentation(pres, "output.pptx")
    """
    presentation.save(path)
