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


def set_slide_background(
    presentation: Presentation,
    slide_number: int,
    *,
    color: str | None = None,
    gradient: dict | None = None,
    image: str | None = None,
    transparency: int = 0,
) -> None:
    """Set the background of a slide.

    Supports solid colors, gradients, and images.

    Args:
        presentation: The presentation to modify
        slide_number: The slide number (1-indexed)
        color: Solid background color as hex ("#0066CC") or theme ("accent1")
        gradient: Gradient definition as dict:
            - colors: List of colors (hex or theme)
            - direction: Angle in degrees (0=left, 90=top, 180=right, 270=bottom)
            - type: "linear" (default) or "radial"
        image: Path to background image file
        transparency: Transparency percentage 0-100 (for color/gradient)

    Note: Only one of color, gradient, or image should be specified.

    Example:
        >>> # Solid color background
        >>> set_slide_background(pres, 1, color="#003366")

        >>> # Theme color with transparency
        >>> set_slide_background(pres, 2, color="accent1", transparency=20)

        >>> # Gradient background
        >>> set_slide_background(pres, 3, gradient={
        ...     "colors": ["#000066", "#0066CC"],
        ...     "direction": 90,  # Top to bottom
        ... })

        >>> # Image background
        >>> set_slide_background(pres, 4, image="background.jpg")
    """
    from pathlib import Path

    from lxml import etree

    from ..oxml.ns import CONTENT_TYPE, REL_TYPE, qn
    from ..oxml.slide import update_slide_in_package

    slide = presentation.get_slide(slide_number)
    slide_part = slide._part
    slide_elem = slide_part._element

    # Find or create cSld/bg element
    c_sld = slide_elem.find(qn("p:cSld"))
    if c_sld is None:
        c_sld = etree.SubElement(slide_elem, qn("p:cSld"))

    # Remove existing background
    existing_bg = c_sld.find(qn("p:bg"))
    if existing_bg is not None:
        c_sld.remove(existing_bg)

    # Create new background element
    bg = etree.Element(qn("p:bg"))
    bg_pr = etree.SubElement(bg, qn("p:bgPr"))

    if color:
        # Solid fill
        solid_fill = etree.SubElement(bg_pr, qn("a:solidFill"))

        if color.startswith("#"):
            # Hex color
            srgb = etree.SubElement(solid_fill, qn("a:srgbClr"))
            srgb.set("val", color.lstrip("#").upper())
            if transparency > 0:
                alpha = etree.SubElement(srgb, qn("a:alpha"))
                alpha.set("val", str((100 - transparency) * 1000))
        else:
            # Theme color
            scheme = etree.SubElement(solid_fill, qn("a:schemeClr"))
            scheme.set("val", color)
            if transparency > 0:
                alpha = etree.SubElement(scheme, qn("a:alpha"))
                alpha.set("val", str((100 - transparency) * 1000))

    elif gradient:
        # Gradient fill
        grad_fill = etree.SubElement(bg_pr, qn("a:gradFill"))

        # Gradient stop list
        gs_lst = etree.SubElement(grad_fill, qn("a:gsLst"))
        colors = gradient.get("colors", ["#FFFFFF", "#000000"])

        for i, grad_color in enumerate(colors):
            pos = int(100000 * i / (len(colors) - 1)) if len(colors) > 1 else 0
            gs = etree.SubElement(gs_lst, qn("a:gs"))
            gs.set("pos", str(pos))

            if grad_color.startswith("#"):
                srgb = etree.SubElement(gs, qn("a:srgbClr"))
                srgb.set("val", grad_color.lstrip("#").upper())
            else:
                scheme = etree.SubElement(gs, qn("a:schemeClr"))
                scheme.set("val", grad_color)

        # Linear gradient
        direction = gradient.get("direction", 90)
        grad_type = gradient.get("type", "linear")

        if grad_type == "linear":
            lin = etree.SubElement(grad_fill, qn("a:lin"))
            # Convert degrees to 60000ths of a degree
            angle = direction * 60000
            lin.set("ang", str(angle))
            lin.set("scaled", "1")
        else:
            # Radial gradient
            path = etree.SubElement(grad_fill, qn("a:path"))
            path.set("path", "circle")
            fill_rect = etree.SubElement(path, qn("a:fillToRect"))
            fill_rect.set("l", "50000")
            fill_rect.set("t", "50000")
            fill_rect.set("r", "50000")
            fill_rect.set("b", "50000")

    elif image:
        # Image fill
        image_path = Path(image)
        if not image_path.exists():
            raise FileNotFoundError(f"Background image not found: {image}")

        # Read image
        with open(image_path, "rb") as f:
            image_data = f.read()

        # Determine content type
        ext = image_path.suffix.lower()
        content_types = {
            ".png": CONTENT_TYPE.PNG,
            ".jpg": CONTENT_TYPE.JPEG,
            ".jpeg": CONTENT_TYPE.JPEG,
            ".gif": CONTENT_TYPE.GIF,
        }
        content_type = content_types.get(ext, CONTENT_TYPE.PNG)

        # Add image to package
        pkg = presentation._package
        existing_images = [
            name for name, _ in pkg.iter_parts()
            if name.startswith("ppt/media/image")
        ]
        image_num = len(existing_images) + 1
        image_part_name = f"ppt/media/image{image_num}{ext}"
        pkg.set_part(image_part_name, image_data, content_type)

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

        # Create blip fill
        blip_fill = etree.SubElement(bg_pr, qn("a:blipFill"))
        blip_fill.set("dpi", "0")
        blip_fill.set("rotWithShape", "1")

        blip = etree.SubElement(blip_fill, qn("a:blip"))
        blip.set(qn("r:embed"), r_id)

        # Stretch
        stretch = etree.SubElement(blip_fill, qn("a:stretch"))
        fill_rect = etree.SubElement(stretch, qn("a:fillRect"))

    else:
        # No fill specified, use noFill
        etree.SubElement(bg_pr, qn("a:noFill"))

    # Effect list (empty)
    etree.SubElement(bg_pr, qn("a:effectLst"))

    # Insert background at the beginning of cSld
    c_sld.insert(0, bg)

    # Save slide
    update_slide_in_package(
        presentation._package,
        slide_number,
        slide_part,
    )


def clear_slide_background(
    presentation: Presentation,
    slide_number: int,
) -> None:
    """Remove custom background from a slide.

    The slide will revert to using the layout/master background.

    Args:
        presentation: The presentation to modify
        slide_number: The slide number (1-indexed)

    Example:
        >>> clear_slide_background(pres, 1)
    """
    from ..oxml.ns import qn
    from ..oxml.slide import update_slide_in_package

    slide = presentation.get_slide(slide_number)
    slide_part = slide._part
    slide_elem = slide_part._element

    # Find and remove background element
    c_sld = slide_elem.find(qn("p:cSld"))
    if c_sld is not None:
        bg = c_sld.find(qn("p:bg"))
        if bg is not None:
            c_sld.remove(bg)

    # Save slide
    update_slide_in_package(
        presentation._package,
        slide_number,
        slide_part,
    )
