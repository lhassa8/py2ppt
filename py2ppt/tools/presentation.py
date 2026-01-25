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


# Slide size presets in EMUs (width, height)
SLIDE_SIZE_PRESETS = {
    "widescreen": (12192000, 6858000),      # 16:9 (13.333" x 7.5")
    "standard": (9144000, 6858000),          # 4:3 (10" x 7.5")
    "widescreen_16x10": (12192000, 7620000), # 16:10 (13.333" x 8.333")
    "a4": (10692000, 7560000),               # A4 landscape
    "a4_portrait": (7560000, 10692000),      # A4 portrait
    "letter": (10058400, 7772400),           # US Letter landscape
    "letter_portrait": (7772400, 10058400),  # US Letter portrait
    "banner": (18288000, 6858000),           # 2:1 banner
    "custom": None,                           # Use explicit dimensions
}


def get_slide_size(presentation: Presentation) -> dict:
    """Get the current slide size.

    Returns:
        Dict with slide dimensions:
        - width: Width in EMUs
        - height: Height in EMUs
        - width_inches: Width in inches
        - height_inches: Height in inches
        - aspect_ratio: Aspect ratio as string (e.g., "16:9")

    Example:
        >>> size = get_slide_size(pres)
        >>> print(f"Size: {size['width_inches']}\" x {size['height_inches']}\"")
    """
    width = presentation.slide_width
    height = presentation.slide_height

    # Convert to inches (914400 EMUs = 1 inch)
    width_inches = round(width / 914400, 2)
    height_inches = round(height / 914400, 2)

    # Determine aspect ratio
    from math import gcd
    divisor = gcd(width, height)
    ratio_w = width // divisor
    ratio_h = height // divisor

    # Simplify common ratios
    if (ratio_w, ratio_h) == (16, 9) or abs(width/height - 16/9) < 0.01:
        aspect_ratio = "16:9"
    elif (ratio_w, ratio_h) == (4, 3) or abs(width/height - 4/3) < 0.01:
        aspect_ratio = "4:3"
    elif (ratio_w, ratio_h) == (16, 10) or abs(width/height - 16/10) < 0.01:
        aspect_ratio = "16:10"
    else:
        aspect_ratio = f"{ratio_w}:{ratio_h}"

    return {
        "width": width,
        "height": height,
        "width_inches": width_inches,
        "height_inches": height_inches,
        "aspect_ratio": aspect_ratio,
    }


def set_slide_size(
    presentation: Presentation,
    width: str | int,
    height: str | int,
) -> None:
    """Set custom slide dimensions.

    Args:
        presentation: The presentation to modify
        width: Slide width (e.g., "10in", "25.4cm", or EMU value)
        height: Slide height

    Example:
        >>> set_slide_size(pres, "13.333in", "7.5in")  # Widescreen
        >>> set_slide_size(pres, "10in", "7.5in")       # Standard 4:3
    """
    from lxml import etree

    from ..oxml.ns import qn
    from ..utils.units import parse_length

    width_emu = int(parse_length(width))
    height_emu = int(parse_length(height))

    # Update presentation part
    pres_part = presentation._presentation
    pres_elem = pres_part._element

    # Find or create sldSz element
    sld_sz = pres_elem.find(qn("p:sldSz"))
    if sld_sz is None:
        # Insert after notesSz if it exists, otherwise at appropriate position
        notes_sz = pres_elem.find(qn("p:notesSz"))
        if notes_sz is not None:
            idx = list(pres_elem).index(notes_sz)
            sld_sz = etree.Element(qn("p:sldSz"))
            pres_elem.insert(idx, sld_sz)
        else:
            sld_sz = etree.SubElement(pres_elem, qn("p:sldSz"))

    sld_sz.set("cx", str(width_emu))
    sld_sz.set("cy", str(height_emu))

    # Update cached dimensions
    presentation._slide_width = width_emu
    presentation._slide_height = height_emu

    # Save presentation part
    from ..oxml.ns import CONTENT_TYPE
    presentation._package.set_part(
        "ppt/presentation.xml",
        pres_part.to_xml(),
        CONTENT_TYPE.PRESENTATION,
    )
    presentation._dirty = True


def set_slide_size_preset(
    presentation: Presentation,
    preset: str,
) -> None:
    """Set slide size to a predefined preset.

    Args:
        presentation: The presentation to modify
        preset: Preset name. Options:
            - "widescreen": 16:9 (default for modern presentations)
            - "standard": 4:3 (traditional ratio)
            - "widescreen_16x10": 16:10
            - "a4": A4 landscape
            - "a4_portrait": A4 portrait
            - "letter": US Letter landscape
            - "letter_portrait": US Letter portrait
            - "banner": 2:1 wide banner

    Example:
        >>> set_slide_size_preset(pres, "widescreen")  # 16:9
        >>> set_slide_size_preset(pres, "standard")    # 4:3
    """
    preset = preset.lower().replace("-", "_").replace(" ", "_")

    if preset not in SLIDE_SIZE_PRESETS:
        available = ", ".join(k for k in SLIDE_SIZE_PRESETS if k != "custom")
        raise ValueError(f"Unknown preset: {preset}. Available: {available}")

    dimensions = SLIDE_SIZE_PRESETS[preset]
    if dimensions is None:
        raise ValueError("Use set_slide_size() for custom dimensions")

    width, height = dimensions
    set_slide_size(presentation, width, height)
