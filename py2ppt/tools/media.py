"""Media tool functions.

Functions for adding tables, images, and other media to slides.
"""

from __future__ import annotations

import os
from pathlib import Path
from typing import Any, List, Optional, Union

from ..core.presentation import Presentation
from ..oxml.shapes import Table, Position
from ..oxml.ns import CONTENT_TYPE, REL_TYPE
from ..utils.units import parse_length


def add_table(
    presentation: Presentation,
    slide_number: int,
    data: List[List[Any]],
    *,
    placeholder: Optional[str] = None,
    left: Optional[Union[str, int]] = None,
    top: Optional[Union[str, int]] = None,
    width: Optional[Union[str, int]] = None,
    height: Optional[Union[str, int]] = None,
) -> None:
    """Add a table to a slide.

    Args:
        presentation: The presentation to modify
        slide_number: The slide number (1-indexed)
        data: 2D list of cell values. First row is typically headers.
        placeholder: Placeholder to fill (e.g., "content", "body")
                    Alternative to specifying position.
        left: Left position (e.g., "1in", "2.5cm")
        top: Top position
        width: Table width
        height: Table height

    Example:
        >>> add_table(pres, 4, data=[
        ...     ["Region", "Q3", "Q4"],
        ...     ["North", 100, 120],
        ...     ["South", 80, 95],
        ... ])
        >>> # Or with explicit position:
        >>> add_table(pres, 4, data=[...],
        ...     left="1in", top="2in", width="8in", height="3in")
    """
    slide = presentation.get_slide(slide_number)
    slide.add_table(
        data,
        left=left,
        top=top,
        width=width,
        height=height,
        placeholder=placeholder,
    )


def update_table_cell(
    presentation: Presentation,
    slide_number: int,
    table_index: int,
    row: int,
    col: int,
    value: Any,
) -> None:
    """Update a single cell in a table.

    Args:
        presentation: The presentation to modify
        slide_number: The slide number (1-indexed)
        table_index: Index of the table on the slide (0-indexed)
        row: Row index (0-indexed)
        col: Column index (0-indexed)
        value: New cell value

    Example:
        >>> update_table_cell(pres, 4, table_index=0, row=1, col=2, value=125)
    """
    slide = presentation.get_slide(slide_number)

    # Find tables on the slide
    tables = [s for s in slide.shapes if isinstance(s, Table)]

    if table_index >= len(tables):
        raise ValueError(
            f"Table index {table_index} out of range. "
            f"Slide has {len(tables)} table(s)."
        )

    table = tables[table_index]

    if row >= len(table.rows):
        raise ValueError(
            f"Row {row} out of range. Table has {len(table.rows)} rows."
        )

    if col >= len(table.rows[row]):
        raise ValueError(
            f"Column {col} out of range. Table has {len(table.rows[row])} columns."
        )

    # Update the cell
    table.rows[row][col].text = str(value) if value is not None else ""

    # Save changes
    from ..oxml.slide import update_slide_in_package

    update_slide_in_package(
        presentation._package,
        slide_number,
        slide._part,
    )


def add_image(
    presentation: Presentation,
    slide_number: int,
    image_path: Union[str, Path],
    *,
    placeholder: Optional[str] = None,
    left: Optional[Union[str, int]] = None,
    top: Optional[Union[str, int]] = None,
    width: Optional[Union[str, int]] = None,
    height: Optional[Union[str, int]] = None,
) -> None:
    """Add an image to a slide.

    Args:
        presentation: The presentation to modify
        slide_number: The slide number (1-indexed)
        image_path: Path to the image file (PNG, JPEG, GIF, etc.)
        placeholder: Placeholder to fill (e.g., "content", "picture")
                    Alternative to specifying position.
        left: Left position (e.g., "1in", "2.5cm")
        top: Top position
        width: Image width. If only width or height is specified,
               the image will scale proportionally.
        height: Image height

    Example:
        >>> add_image(pres, 3, "chart.png", placeholder="content")
        >>> add_image(pres, 3, "logo.png",
        ...     left="7in", top="0.5in", width="2in")
    """
    image_path = Path(image_path)

    if not image_path.exists():
        raise FileNotFoundError(f"Image file not found: {image_path}")

    # Read image data
    with open(image_path, "rb") as f:
        image_data = f.read()

    # Determine content type
    ext = image_path.suffix.lower()
    content_types = {
        ".png": CONTENT_TYPE.PNG,
        ".jpg": CONTENT_TYPE.JPEG,
        ".jpeg": CONTENT_TYPE.JPEG,
        ".gif": CONTENT_TYPE.GIF,
        ".bmp": CONTENT_TYPE.BMP,
        ".tiff": CONTENT_TYPE.TIFF,
        ".tif": CONTENT_TYPE.TIFF,
        ".emf": CONTENT_TYPE.EMF,
        ".wmf": CONTENT_TYPE.WMF,
    }

    content_type = content_types.get(ext)
    if content_type is None:
        raise ValueError(f"Unsupported image format: {ext}")

    # Add image to package
    pkg = presentation._package

    # Find next available image number
    existing_images = [
        name for name, _ in pkg.iter_parts()
        if name.startswith("ppt/media/image")
    ]
    image_num = len(existing_images) + 1
    image_part_name = f"ppt/media/image{image_num}{ext}"

    pkg.set_part(image_part_name, image_data, content_type)

    # Get slide and determine position
    slide = presentation.get_slide(slide_number)

    if placeholder:
        ph_shape = slide._find_placeholder(placeholder)
        position = ph_shape.position
    elif left is not None and top is not None:
        position = Position(
            x=int(parse_length(left)),
            y=int(parse_length(top)),
            cx=int(parse_length(width or "4in")),
            cy=int(parse_length(height or "3in")),
        )
    else:
        # Default position (centered)
        position = Position(
            x=presentation.slide_width // 4,
            y=presentation.slide_height // 4,
            cx=presentation.slide_width // 2,
            cy=presentation.slide_height // 2,
        )

    # Create picture shape
    from ..oxml.shapes import Picture

    # Add relationship to slide
    slide_refs = presentation._presentation.get_slide_refs()
    slide_ref = slide_refs[slide_number - 1]
    pres_rels = pkg.get_part_rels("ppt/presentation.xml")
    rel = pres_rels.get(slide_ref.r_id)

    if rel.target.startswith("/"):
        slide_path = rel.target.lstrip("/")
    else:
        slide_path = f"ppt/{rel.target}"

    slide_rels = pkg.get_part_rels(slide_path)
    r_id = slide_rels.add(
        rel_type=REL_TYPE.IMAGE,
        target=f"../media/image{image_num}{ext}",
    )
    pkg.set_part_rels(slide_path, slide_rels)

    # Create picture
    pic = Picture(
        id=slide._part.shape_tree._next_id,
        name=f"Picture {image_num}",
        position=position,
        r_embed=r_id,
    )

    slide._part.add_picture(pic)
    slide._save()
