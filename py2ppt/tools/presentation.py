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


# ============================================================================
# Document Properties
# ============================================================================


def set_document_property(
    presentation: Presentation,
    name: str,
    value: str,
) -> None:
    """Set a document property (core metadata).

    Args:
        presentation: The presentation to modify
        name: Property name. Standard properties:
            - "title": Document title
            - "author" or "creator": Document author
            - "subject": Document subject
            - "description": Document description
            - "keywords": Keywords (comma-separated)
            - "category": Document category
            - "last_modified_by": Last editor
        value: Property value

    Example:
        >>> set_document_property(pres, "title", "Q4 Report")
        >>> set_document_property(pres, "author", "John Smith")
        >>> set_document_property(pres, "keywords", "quarterly, finance, report")
    """
    from lxml import etree

    from ..oxml.ns import CONTENT_TYPE, qn

    pkg = presentation._package

    # Get or create core properties
    core_xml = pkg.get_part("docProps/core.xml")

    if core_xml is None:
        # Create minimal core.xml
        core_xml = b'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties"
    xmlns:dc="http://purl.org/dc/elements/1.1/"
    xmlns:dcterms="http://purl.org/dc/terms/"
    xmlns:dcmitype="http://purl.org/dc/dcmitype/"
    xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
</cp:coreProperties>'''

    root = etree.fromstring(core_xml)

    # Map property names to XML elements
    property_map = {
        "title": "dc:title",
        "author": "dc:creator",
        "creator": "dc:creator",
        "subject": "dc:subject",
        "description": "dc:description",
        "keywords": "cp:keywords",
        "category": "cp:category",
        "last_modified_by": "cp:lastModifiedBy",
        "lastmodifiedby": "cp:lastModifiedBy",
    }

    prop_name = name.lower().replace(" ", "_").replace("-", "_")
    xml_tag = property_map.get(prop_name)

    if xml_tag is None:
        raise ValueError(
            f"Unknown property: {name}. "
            f"Available: {', '.join(property_map.keys())}"
        )

    # Find or create the element
    elem = root.find(qn(xml_tag))
    if elem is None:
        elem = etree.SubElement(root, qn(xml_tag))

    elem.text = value

    # Save
    xml_bytes = etree.tostring(
        root,
        xml_declaration=True,
        encoding="UTF-8",
        standalone=True,
    )
    pkg.set_part("docProps/core.xml", xml_bytes, CONTENT_TYPE.CORE_PROPS)
    presentation._dirty = True


def get_document_property(
    presentation: Presentation,
    name: str,
) -> str | None:
    """Get a document property value.

    Args:
        presentation: The presentation to inspect
        name: Property name (see set_document_property for options)

    Returns:
        Property value or None if not set

    Example:
        >>> title = get_document_property(pres, "title")
        >>> author = get_document_property(pres, "author")
    """
    from lxml import etree

    from ..oxml.ns import qn

    pkg = presentation._package
    core_xml = pkg.get_part("docProps/core.xml")

    if core_xml is None:
        return None

    root = etree.fromstring(core_xml)

    property_map = {
        "title": "dc:title",
        "author": "dc:creator",
        "creator": "dc:creator",
        "subject": "dc:subject",
        "description": "dc:description",
        "keywords": "cp:keywords",
        "category": "cp:category",
        "last_modified_by": "cp:lastModifiedBy",
        "lastmodifiedby": "cp:lastModifiedBy",
    }

    prop_name = name.lower().replace(" ", "_").replace("-", "_")
    xml_tag = property_map.get(prop_name)

    if xml_tag is None:
        return None

    elem = root.find(qn(xml_tag))
    return elem.text if elem is not None else None


def get_document_info(presentation: Presentation) -> dict:
    """Get all document properties.

    Returns:
        Dict with all available document metadata

    Example:
        >>> info = get_document_info(pres)
        >>> print(f"Title: {info.get('title')}")
        >>> print(f"Author: {info.get('author')}")
    """
    from lxml import etree

    from ..oxml.ns import qn

    pkg = presentation._package
    info = {}

    # Core properties
    core_xml = pkg.get_part("docProps/core.xml")
    if core_xml is not None:
        root = etree.fromstring(core_xml)

        property_map = {
            "dc:title": "title",
            "dc:creator": "author",
            "dc:subject": "subject",
            "dc:description": "description",
            "cp:keywords": "keywords",
            "cp:category": "category",
            "cp:lastModifiedBy": "last_modified_by",
            "dcterms:created": "created",
            "dcterms:modified": "modified",
        }

        for xml_tag, prop_name in property_map.items():
            elem = root.find(qn(xml_tag))
            if elem is not None and elem.text:
                info[prop_name] = elem.text

    # Extended properties (app.xml)
    app_xml = pkg.get_part("docProps/app.xml")
    if app_xml is not None:
        root = etree.fromstring(app_xml)
        ns = "http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"

        for tag, key in [
            ("Application", "application"),
            ("Company", "company"),
            ("Slides", "slide_count"),
            ("Notes", "notes_count"),
            ("HiddenSlides", "hidden_slides"),
        ]:
            elem = root.find(f"{{{ns}}}{tag}")
            if elem is not None and elem.text:
                info[key] = elem.text

    return info


def set_custom_property(
    presentation: Presentation,
    name: str,
    value: str | int | float | bool,
) -> None:
    """Set a custom document property.

    Custom properties allow storing arbitrary metadata.

    Args:
        presentation: The presentation to modify
        name: Property name (any string)
        value: Property value (string, int, float, or bool)

    Example:
        >>> set_custom_property(pres, "Project Code", "PRJ-2024-001")
        >>> set_custom_property(pres, "Approved", True)
        >>> set_custom_property(pres, "Version", 2.5)
    """
    from lxml import etree

    pkg = presentation._package

    custom_ns = "http://schemas.openxmlformats.org/officeDocument/2006/custom-properties"
    vt_ns = "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"

    # Get or create custom properties
    custom_xml = pkg.get_part("docProps/custom.xml")

    if custom_xml is None:
        # Create new custom.xml
        root = etree.Element(
            f"{{{custom_ns}}}Properties",
            nsmap={None: custom_ns, "vt": vt_ns},
        )
        next_pid = 2  # PIDs start at 2
    else:
        root = etree.fromstring(custom_xml)
        # Find highest PID
        pids = [int(p.get("pid", "1")) for p in root]
        next_pid = max(pids) + 1 if pids else 2

    # Check if property already exists
    existing = None
    for prop in root:
        if prop.get("name") == name:
            existing = prop
            break

    if existing is not None:
        # Update existing
        root.remove(existing)

    # Determine value type
    if isinstance(value, bool):
        vt_type = "bool"
        vt_value = "true" if value else "false"
    elif isinstance(value, int):
        vt_type = "i4"
        vt_value = str(value)
    elif isinstance(value, float):
        vt_type = "r8"
        vt_value = str(value)
    else:
        vt_type = "lpwstr"
        vt_value = str(value)

    # Create property element
    prop = etree.SubElement(
        root,
        f"{{{custom_ns}}}property",
        attrib={
            "fmtid": "{D5CDD505-2E9C-101B-9397-08002B2CF9AE}",
            "pid": str(next_pid if existing is None else existing.get("pid")),
            "name": name,
        },
    )

    value_elem = etree.SubElement(prop, f"{{{vt_ns}}}{vt_type}")
    value_elem.text = vt_value

    # Save
    xml_bytes = etree.tostring(
        root,
        xml_declaration=True,
        encoding="UTF-8",
        standalone=True,
    )

    content_type = "application/vnd.openxmlformats-officedocument.custom-properties+xml"
    pkg.set_part("docProps/custom.xml", xml_bytes, content_type)

    # Add relationship if needed
    # Check if custom props relationship exists
    has_custom_rel = any(
        r.target == "docProps/custom.xml" for r in pkg.package_rels
    )
    if not has_custom_rel:
        pkg.package_rels.add(
            rel_type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/custom-properties",
            target="docProps/custom.xml",
        )

    presentation._dirty = True


def get_custom_property(
    presentation: Presentation,
    name: str,
) -> str | int | float | bool | None:
    """Get a custom document property value.

    Args:
        presentation: The presentation to inspect
        name: Property name

    Returns:
        Property value (typed appropriately) or None if not found

    Example:
        >>> code = get_custom_property(pres, "Project Code")
        >>> approved = get_custom_property(pres, "Approved")
    """
    from lxml import etree

    pkg = presentation._package
    custom_xml = pkg.get_part("docProps/custom.xml")

    if custom_xml is None:
        return None

    root = etree.fromstring(custom_xml)

    for prop in root:
        if prop.get("name") == name:
            # Find value element
            for child in prop:
                tag = child.tag.split("}")[-1] if "}" in child.tag else child.tag
                text = child.text

                if text is None:
                    return None

                if tag == "bool":
                    return text.lower() == "true"
                elif tag == "i4":
                    return int(text)
                elif tag == "r8":
                    return float(text)
                else:
                    return text

    return None


def get_custom_properties(presentation: Presentation) -> dict:
    """Get all custom document properties.

    Returns:
        Dict of property name -> value

    Example:
        >>> props = get_custom_properties(pres)
        >>> for name, value in props.items():
        ...     print(f"{name}: {value}")
    """
    from lxml import etree

    pkg = presentation._package
    custom_xml = pkg.get_part("docProps/custom.xml")

    if custom_xml is None:
        return {}

    root = etree.fromstring(custom_xml)
    props = {}

    for prop in root:
        name = prop.get("name")
        if name is None:
            continue

        # Find value element
        for child in prop:
            tag = child.tag.split("}")[-1] if "}" in child.tag else child.tag
            text = child.text

            if text is None:
                props[name] = None
            elif tag == "bool":
                props[name] = text.lower() == "true"
            elif tag == "i4":
                props[name] = int(text)
            elif tag == "r8":
                props[name] = float(text)
            else:
                props[name] = text
            break

    return props


def remove_custom_property(
    presentation: Presentation,
    name: str,
) -> bool:
    """Remove a custom document property.

    Args:
        presentation: The presentation to modify
        name: Property name to remove

    Returns:
        True if property was removed, False if not found

    Example:
        >>> remove_custom_property(pres, "Project Code")
    """
    from lxml import etree

    pkg = presentation._package
    custom_xml = pkg.get_part("docProps/custom.xml")

    if custom_xml is None:
        return False

    root = etree.fromstring(custom_xml)

    # Find and remove property
    for prop in root:
        if prop.get("name") == name:
            root.remove(prop)

            # Save
            xml_bytes = etree.tostring(
                root,
                xml_declaration=True,
                encoding="UTF-8",
                standalone=True,
            )
            content_type = "application/vnd.openxmlformats-officedocument.custom-properties+xml"
            pkg.set_part("docProps/custom.xml", xml_bytes, content_type)
            presentation._dirty = True
            return True

    return False
