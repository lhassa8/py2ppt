"""Master slide manipulation tool functions.

Functions for inspecting and editing slide masters and layouts.
Masters define the default styling that applies to all slides using them.
"""

from __future__ import annotations

from typing import Any

from ..core.presentation import Presentation


def get_slide_masters(presentation: Presentation) -> list[dict[str, Any]]:
    """Get all slide masters in the presentation.

    Args:
        presentation: The presentation to inspect

    Returns:
        List of master dicts with:
        - index: Master index (0-indexed)
        - name: Master name
        - layouts: List of layout names using this master

    Example:
        >>> masters = get_slide_masters(pres)
        >>> for m in masters:
        ...     print(f"{m['name']}: {len(m['layouts'])} layouts")
    """
    from lxml import etree

    from ..oxml.ns import qn

    pkg = presentation._package
    pres_part = presentation._presentation
    pres_elem = pres_part._element

    masters = []

    # Get slide master references
    sld_master_id_lst = pres_elem.find(qn("p:sldMasterIdLst"))
    if sld_master_id_lst is None:
        return masters

    pres_rels = pkg.get_part_rels("ppt/presentation.xml")

    for idx, sld_master_id in enumerate(sld_master_id_lst.findall(qn("p:sldMasterId"))):
        r_id = sld_master_id.get(qn("r:id"))
        rel = pres_rels.get(r_id)

        if rel is None:
            continue

        # Get master path
        if rel.target.startswith("/"):
            master_path = rel.target.lstrip("/")
        else:
            master_path = f"ppt/{rel.target}"

        # Read master XML
        master_xml = pkg.get_part(master_path)
        if master_xml is None:
            continue

        master_elem = etree.fromstring(master_xml)

        # Get master name from cSld
        c_sld = master_elem.find(qn("p:cSld"))
        master_name = c_sld.get("name", f"Master {idx + 1}") if c_sld is not None else f"Master {idx + 1}"

        # Get layouts for this master
        layouts = []
        master_rels = pkg.get_part_rels(master_path)

        from ..oxml.ns import REL_TYPE
        layout_rels = master_rels.find_by_type(REL_TYPE.SLIDE_LAYOUT)

        for layout_rel in layout_rels:
            # Get layout path
            if layout_rel.target.startswith("/"):
                layout_path = layout_rel.target.lstrip("/")
            else:
                # Relative to master
                layout_path = f"ppt/slideLayouts/{layout_rel.target.split('/')[-1]}"

            layout_xml = pkg.get_part(layout_path)
            if layout_xml is None:
                continue

            layout_elem = etree.fromstring(layout_xml)
            layout_c_sld = layout_elem.find(qn("p:cSld"))
            layout_name = layout_c_sld.get("name", "Unnamed") if layout_c_sld is not None else "Unnamed"
            layouts.append(layout_name)

        masters.append({
            "index": idx,
            "name": master_name,
            "layouts": layouts,
        })

    return masters


def get_master_placeholders(
    presentation: Presentation,
    master_index: int = 0,
) -> list[dict[str, Any]]:
    """Get placeholders defined in a slide master.

    Args:
        presentation: The presentation to inspect
        master_index: Master index (0-indexed, default 0)

    Returns:
        List of placeholder dicts with:
        - type: Placeholder type (title, body, dt, ftr, sldNum)
        - idx: Placeholder index
        - position: Dict with left, top, width, height in EMUs

    Example:
        >>> placeholders = get_master_placeholders(pres)
        >>> for ph in placeholders:
        ...     print(f"{ph['type']}: {ph['position']}")
    """
    from lxml import etree

    from ..oxml.ns import qn

    pkg = presentation._package
    pres_part = presentation._presentation
    pres_elem = pres_part._element

    # Get slide master references
    sld_master_id_lst = pres_elem.find(qn("p:sldMasterIdLst"))
    if sld_master_id_lst is None:
        return []

    master_ids = sld_master_id_lst.findall(qn("p:sldMasterId"))
    if master_index < 0 or master_index >= len(master_ids):
        return []

    r_id = master_ids[master_index].get(qn("r:id"))
    pres_rels = pkg.get_part_rels("ppt/presentation.xml")
    rel = pres_rels.get(r_id)

    if rel is None:
        return []

    # Get master path
    if rel.target.startswith("/"):
        master_path = rel.target.lstrip("/")
    else:
        master_path = f"ppt/{rel.target}"

    # Read master XML
    master_xml = pkg.get_part(master_path)
    if master_xml is None:
        return []

    master_elem = etree.fromstring(master_xml)

    placeholders = []

    # Find shape tree
    c_sld = master_elem.find(qn("p:cSld"))
    if c_sld is None:
        return []

    sp_tree = c_sld.find(qn("p:spTree"))
    if sp_tree is None:
        return []

    # Look for shapes with placeholders
    for sp in sp_tree.findall(qn("p:sp")):
        nv_sp_pr = sp.find(qn("p:nvSpPr"))
        if nv_sp_pr is None:
            continue

        nv_pr = nv_sp_pr.find(qn("p:nvPr"))
        if nv_pr is None:
            continue

        ph = nv_pr.find(qn("p:ph"))
        if ph is None:
            continue

        ph_type = ph.get("type", "body")
        ph_idx = ph.get("idx", "0")

        # Get position
        sp_pr = sp.find(qn("p:spPr"))
        position = {"left": 0, "top": 0, "width": 0, "height": 0}

        if sp_pr is not None:
            xfrm = sp_pr.find(qn("a:xfrm"))
            if xfrm is not None:
                off = xfrm.find(qn("a:off"))
                if off is not None:
                    position["left"] = int(off.get("x", 0))
                    position["top"] = int(off.get("y", 0))

                ext = xfrm.find(qn("a:ext"))
                if ext is not None:
                    position["width"] = int(ext.get("cx", 0))
                    position["height"] = int(ext.get("cy", 0))

        placeholders.append({
            "type": ph_type,
            "idx": int(ph_idx) if ph_idx.isdigit() else 0,
            "position": position,
        })

    return placeholders


def set_master_background(
    presentation: Presentation,
    master_index: int = 0,
    *,
    color: str | None = None,
    gradient: dict | None = None,
    image: str | None = None,
    transparency: int = 0,
) -> bool:
    """Set the background of a slide master.

    Changes apply to all slides using this master (unless overridden).

    Args:
        presentation: The presentation to modify
        master_index: Master index (0-indexed, default 0)
        color: Solid background color as hex ("#0066CC") or theme ("accent1")
        gradient: Gradient definition as dict:
            - colors: List of colors (hex or theme)
            - direction: Angle in degrees (0=left, 90=top)
        image: Path to background image file
        transparency: Transparency percentage 0-100

    Returns:
        True if successful, False otherwise

    Example:
        >>> set_master_background(pres, color="#003366")
        >>> set_master_background(pres, gradient={"colors": ["#000066", "#0066CC"], "direction": 90})
    """
    from pathlib import Path

    from lxml import etree

    from ..oxml.ns import CONTENT_TYPE, REL_TYPE, qn

    pkg = presentation._package
    pres_part = presentation._presentation
    pres_elem = pres_part._element

    # Get slide master references
    sld_master_id_lst = pres_elem.find(qn("p:sldMasterIdLst"))
    if sld_master_id_lst is None:
        return False

    master_ids = sld_master_id_lst.findall(qn("p:sldMasterId"))
    if master_index < 0 or master_index >= len(master_ids):
        return False

    r_id = master_ids[master_index].get(qn("r:id"))
    pres_rels = pkg.get_part_rels("ppt/presentation.xml")
    rel = pres_rels.get(r_id)

    if rel is None:
        return False

    # Get master path
    if rel.target.startswith("/"):
        master_path = rel.target.lstrip("/")
    else:
        master_path = f"ppt/{rel.target}"

    # Read master XML
    master_xml = pkg.get_part(master_path)
    if master_xml is None:
        return False

    master_elem = etree.fromstring(master_xml)

    # Find or create cSld/bg element
    c_sld = master_elem.find(qn("p:cSld"))
    if c_sld is None:
        return False

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
            srgb = etree.SubElement(solid_fill, qn("a:srgbClr"))
            srgb.set("val", color.lstrip("#").upper())
            if transparency > 0:
                alpha = etree.SubElement(srgb, qn("a:alpha"))
                alpha.set("val", str((100 - transparency) * 1000))
        else:
            scheme = etree.SubElement(solid_fill, qn("a:schemeClr"))
            scheme.set("val", color)
            if transparency > 0:
                alpha = etree.SubElement(scheme, qn("a:alpha"))
                alpha.set("val", str((100 - transparency) * 1000))

    elif gradient:
        # Gradient fill
        grad_fill = etree.SubElement(bg_pr, qn("a:gradFill"))
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

        direction = gradient.get("direction", 90)
        lin = etree.SubElement(grad_fill, qn("a:lin"))
        lin.set("ang", str(direction * 60000))
        lin.set("scaled", "1")

    elif image:
        # Image fill
        image_path = Path(image)
        if not image_path.exists():
            return False

        with open(image_path, "rb") as f:
            image_data = f.read()

        ext = image_path.suffix.lower()
        content_types = {
            ".png": CONTENT_TYPE.PNG,
            ".jpg": CONTENT_TYPE.JPEG,
            ".jpeg": CONTENT_TYPE.JPEG,
            ".gif": CONTENT_TYPE.GIF,
        }
        content_type = content_types.get(ext, CONTENT_TYPE.PNG)

        # Add image to package
        existing_images = [
            name for name, _ in pkg.iter_parts()
            if name.startswith("ppt/media/image")
        ]
        image_num = len(existing_images) + 1
        image_part_name = f"ppt/media/image{image_num}{ext}"
        pkg.set_part(image_part_name, image_data, content_type)

        # Add relationship to master
        master_rels = pkg.get_part_rels(master_path)
        img_r_id = master_rels.add(
            rel_type=REL_TYPE.IMAGE,
            target=f"../media/image{image_num}{ext}",
        )
        pkg.set_part_rels(master_path, master_rels)

        # Create blip fill
        blip_fill = etree.SubElement(bg_pr, qn("a:blipFill"))
        blip_fill.set("dpi", "0")
        blip_fill.set("rotWithShape", "1")

        blip = etree.SubElement(blip_fill, qn("a:blip"))
        blip.set(qn("r:embed"), img_r_id)

        stretch = etree.SubElement(blip_fill, qn("a:stretch"))
        etree.SubElement(stretch, qn("a:fillRect"))

    else:
        etree.SubElement(bg_pr, qn("a:noFill"))

    etree.SubElement(bg_pr, qn("a:effectLst"))

    # Insert background at beginning of cSld
    c_sld.insert(0, bg)

    # Save master
    xml_bytes = etree.tostring(
        master_elem,
        xml_declaration=True,
        encoding="UTF-8",
        standalone=True,
    )

    # Get content type for master
    master_content_type = "application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml"
    pkg.set_part(master_path, xml_bytes, master_content_type)
    presentation._dirty = True

    return True


def set_master_font(
    presentation: Presentation,
    master_index: int = 0,
    placeholder_type: str = "title",
    *,
    font_name: str | None = None,
    font_size: int | None = None,
    color: str | None = None,
    bold: bool | None = None,
    italic: bool | None = None,
) -> bool:
    """Set font properties for a placeholder type in the master.

    Changes apply to all slides using this master for the specified placeholder.

    Args:
        presentation: The presentation to modify
        master_index: Master index (0-indexed, default 0)
        placeholder_type: Type of placeholder to style ("title", "body", "dt", "ftr", "sldNum")
        font_name: Font family name
        font_size: Font size in points
        color: Font color as hex or theme color
        bold: Whether text should be bold
        italic: Whether text should be italic

    Returns:
        True if successful, False otherwise

    Example:
        >>> set_master_font(pres, placeholder_type="title", font_name="Arial", font_size=44)
        >>> set_master_font(pres, placeholder_type="body", color="#333333")
    """
    from lxml import etree

    from ..oxml.ns import qn

    pkg = presentation._package
    pres_part = presentation._presentation
    pres_elem = pres_part._element

    # Get slide master references
    sld_master_id_lst = pres_elem.find(qn("p:sldMasterIdLst"))
    if sld_master_id_lst is None:
        return False

    master_ids = sld_master_id_lst.findall(qn("p:sldMasterId"))
    if master_index < 0 or master_index >= len(master_ids):
        return False

    r_id = master_ids[master_index].get(qn("r:id"))
    pres_rels = pkg.get_part_rels("ppt/presentation.xml")
    rel = pres_rels.get(r_id)

    if rel is None:
        return False

    # Get master path
    if rel.target.startswith("/"):
        master_path = rel.target.lstrip("/")
    else:
        master_path = f"ppt/{rel.target}"

    # Read master XML
    master_xml = pkg.get_part(master_path)
    if master_xml is None:
        return False

    master_elem = etree.fromstring(master_xml)

    # Find the placeholder shape
    c_sld = master_elem.find(qn("p:cSld"))
    if c_sld is None:
        return False

    sp_tree = c_sld.find(qn("p:spTree"))
    if sp_tree is None:
        return False

    target_shape = None
    for sp in sp_tree.findall(qn("p:sp")):
        nv_sp_pr = sp.find(qn("p:nvSpPr"))
        if nv_sp_pr is None:
            continue

        nv_pr = nv_sp_pr.find(qn("p:nvPr"))
        if nv_pr is None:
            continue

        ph = nv_pr.find(qn("p:ph"))
        if ph is None:
            continue

        ph_type = ph.get("type", "body")
        if ph_type == placeholder_type:
            target_shape = sp
            break

    if target_shape is None:
        return False

    # Find or create txBody
    tx_body = target_shape.find(qn("p:txBody"))
    if tx_body is None:
        return False

    # Get the first paragraph (or create default properties)
    lst_style = tx_body.find(qn("a:lstStyle"))

    # We'll modify the first paragraph's default run properties
    # For title placeholder, look for lvl1pPr
    if lst_style is None:
        lst_style = etree.SubElement(tx_body, qn("a:lstStyle"))

    lvl1_pr = lst_style.find(qn("a:lvl1pPr"))
    if lvl1_pr is None:
        lvl1_pr = etree.SubElement(lst_style, qn("a:lvl1pPr"))

    def_rpr = lvl1_pr.find(qn("a:defRPr"))
    if def_rpr is None:
        def_rpr = etree.SubElement(lvl1_pr, qn("a:defRPr"))

    # Apply font properties
    if font_size is not None:
        def_rpr.set("sz", str(font_size * 100))

    if bold is not None:
        def_rpr.set("b", "1" if bold else "0")

    if italic is not None:
        def_rpr.set("i", "1" if italic else "0")

    if font_name is not None:
        # Remove existing font settings
        for latin in def_rpr.findall(qn("a:latin")):
            def_rpr.remove(latin)

        latin = etree.SubElement(def_rpr, qn("a:latin"))
        latin.set("typeface", font_name)

    if color is not None:
        # Remove existing color
        for solid_fill in def_rpr.findall(qn("a:solidFill")):
            def_rpr.remove(solid_fill)

        solid_fill = etree.SubElement(def_rpr, qn("a:solidFill"))
        if color.startswith("#"):
            srgb = etree.SubElement(solid_fill, qn("a:srgbClr"))
            srgb.set("val", color.lstrip("#").upper())
        else:
            scheme = etree.SubElement(solid_fill, qn("a:schemeClr"))
            scheme.set("val", color)

    # Save master
    xml_bytes = etree.tostring(
        master_elem,
        xml_declaration=True,
        encoding="UTF-8",
        standalone=True,
    )

    master_content_type = "application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml"
    pkg.set_part(master_path, xml_bytes, master_content_type)
    presentation._dirty = True

    return True


def get_layout_info(
    presentation: Presentation,
    layout_name: str | None = None,
    layout_index: int | None = None,
) -> dict[str, Any] | None:
    """Get detailed information about a layout.

    Args:
        presentation: The presentation to inspect
        layout_name: Layout name to find (fuzzy matched)
        layout_index: Layout index (0-indexed)

    Returns:
        Layout info dict with:
        - name: Layout name
        - index: Layout index
        - master_index: Parent master index
        - placeholders: List of placeholder types

    Example:
        >>> info = get_layout_info(pres, layout_name="Title Slide")
        >>> print(f"Placeholders: {info['placeholders']}")
    """
    from lxml import etree

    from ..oxml.ns import REL_TYPE, qn

    pkg = presentation._package
    pres_part = presentation._presentation
    pres_elem = pres_part._element

    # Get all layouts
    layouts = presentation.get_layouts()

    target_layout = None
    if layout_name is not None:
        # Fuzzy match
        layout_name_lower = layout_name.lower()
        for layout in layouts:
            if layout.name.lower() == layout_name_lower or layout_name_lower in layout.name.lower():
                target_layout = layout
                break
    elif layout_index is not None:
        if 0 <= layout_index < len(layouts):
            target_layout = layouts[layout_index]

    if target_layout is None:
        return None

    # Get layout XML to find placeholders
    sld_master_id_lst = pres_elem.find(qn("p:sldMasterIdLst"))
    if sld_master_id_lst is None:
        return None

    pres_rels = pkg.get_part_rels("ppt/presentation.xml")

    placeholders = []
    master_index = 0

    # Search through masters to find the layout
    for m_idx, sld_master_id in enumerate(sld_master_id_lst.findall(qn("p:sldMasterId"))):
        r_id = sld_master_id.get(qn("r:id"))
        rel = pres_rels.get(r_id)

        if rel is None:
            continue

        if rel.target.startswith("/"):
            master_path = rel.target.lstrip("/")
        else:
            master_path = f"ppt/{rel.target}"

        master_rels = pkg.get_part_rels(master_path)
        layout_rels = master_rels.find_by_type(REL_TYPE.SLIDE_LAYOUT)

        for layout_rel in layout_rels:
            if layout_rel.target.startswith("/"):
                layout_path = layout_rel.target.lstrip("/")
            else:
                layout_path = f"ppt/slideLayouts/{layout_rel.target.split('/')[-1]}"

            layout_xml = pkg.get_part(layout_path)
            if layout_xml is None:
                continue

            layout_elem = etree.fromstring(layout_xml)
            layout_c_sld = layout_elem.find(qn("p:cSld"))
            current_name = layout_c_sld.get("name", "") if layout_c_sld is not None else ""

            if current_name == target_layout.name:
                master_index = m_idx

                # Get placeholders
                if layout_c_sld is not None:
                    sp_tree = layout_c_sld.find(qn("p:spTree"))
                    if sp_tree is not None:
                        for sp in sp_tree.findall(qn("p:sp")):
                            nv_sp_pr = sp.find(qn("p:nvSpPr"))
                            if nv_sp_pr is None:
                                continue
                            nv_pr = nv_sp_pr.find(qn("p:nvPr"))
                            if nv_pr is None:
                                continue
                            ph = nv_pr.find(qn("p:ph"))
                            if ph is not None:
                                placeholders.append(ph.get("type", "body"))
                break

    return {
        "name": target_layout.name,
        "index": target_layout.index,
        "master_index": master_index,
        "placeholders": placeholders,
    }


def add_logo_to_master(
    presentation: Presentation,
    image_path: str,
    *,
    master_index: int = 0,
    left: str | int = "0.5in",
    top: str | int = "0.3in",
    width: str | int = "1in",
    height: str | int | None = None,
) -> bool:
    """Add a logo image to the slide master.

    The logo will appear on all slides using this master.

    Args:
        presentation: The presentation to modify
        image_path: Path to the logo image file
        master_index: Master index (0-indexed, default 0)
        left: Left position
        top: Top position
        width: Width of the logo
        height: Height (if None, maintains aspect ratio based on width)

    Returns:
        True if successful, False otherwise

    Example:
        >>> add_logo_to_master(pres, "logo.png", left="0.5in", top="0.3in", width="1.5in")
    """
    from pathlib import Path

    from lxml import etree

    from ..oxml.ns import CONTENT_TYPE, REL_TYPE, qn
    from ..utils.units import parse_length

    image_file = Path(image_path)
    if not image_file.exists():
        return False

    pkg = presentation._package
    pres_part = presentation._presentation
    pres_elem = pres_part._element

    # Get slide master references
    sld_master_id_lst = pres_elem.find(qn("p:sldMasterIdLst"))
    if sld_master_id_lst is None:
        return False

    master_ids = sld_master_id_lst.findall(qn("p:sldMasterId"))
    if master_index < 0 or master_index >= len(master_ids):
        return False

    r_id = master_ids[master_index].get(qn("r:id"))
    pres_rels = pkg.get_part_rels("ppt/presentation.xml")
    rel = pres_rels.get(r_id)

    if rel is None:
        return False

    # Get master path
    if rel.target.startswith("/"):
        master_path = rel.target.lstrip("/")
    else:
        master_path = f"ppt/{rel.target}"

    # Read master XML
    master_xml = pkg.get_part(master_path)
    if master_xml is None:
        return False

    master_elem = etree.fromstring(master_xml)

    # Read image file
    with open(image_file, "rb") as f:
        image_data = f.read()

    # Determine content type
    ext = image_file.suffix.lower()
    content_types = {
        ".png": CONTENT_TYPE.PNG,
        ".jpg": CONTENT_TYPE.JPEG,
        ".jpeg": CONTENT_TYPE.JPEG,
        ".gif": CONTENT_TYPE.GIF,
    }
    content_type = content_types.get(ext, CONTENT_TYPE.PNG)

    # Add image to package
    existing_images = [
        name for name, _ in pkg.iter_parts()
        if name.startswith("ppt/media/image")
    ]
    image_num = len(existing_images) + 1
    image_part_name = f"ppt/media/image{image_num}{ext}"
    pkg.set_part(image_part_name, image_data, content_type)

    # Add relationship to master
    master_rels = pkg.get_part_rels(master_path)
    img_r_id = master_rels.add(
        rel_type=REL_TYPE.IMAGE,
        target=f"../media/image{image_num}{ext}",
    )
    pkg.set_part_rels(master_path, master_rels)

    # Parse dimensions
    x = int(parse_length(left)) if isinstance(left, str) else left
    y = int(parse_length(top)) if isinstance(top, str) else top
    cx = int(parse_length(width)) if isinstance(width, str) else width

    # If height not specified, use same as width (square)
    if height is None:
        cy = cx
    else:
        cy = int(parse_length(height)) if isinstance(height, str) else height

    # Find shape tree
    c_sld = master_elem.find(qn("p:cSld"))
    if c_sld is None:
        return False

    sp_tree = c_sld.find(qn("p:spTree"))
    if sp_tree is None:
        return False

    # Get next shape ID
    max_id = 1
    for sp in sp_tree.iter():
        c_nv_pr = sp.find(qn("p:cNvPr"))
        if c_nv_pr is not None:
            try:
                shape_id = int(c_nv_pr.get("id", 0))
                max_id = max(max_id, shape_id)
            except ValueError:
                pass
    next_id = max_id + 1

    # Create picture element
    pic = etree.SubElement(sp_tree, qn("p:pic"))

    # Non-visual properties
    nv_pic_pr = etree.SubElement(pic, qn("p:nvPicPr"))
    c_nv_pr = etree.SubElement(nv_pic_pr, qn("p:cNvPr"))
    c_nv_pr.set("id", str(next_id))
    c_nv_pr.set("name", f"Logo {next_id}")

    c_nv_pic_pr = etree.SubElement(nv_pic_pr, qn("p:cNvPicPr"))
    pic_locks = etree.SubElement(c_nv_pic_pr, qn("a:picLocks"))
    pic_locks.set("noChangeAspect", "1")

    etree.SubElement(nv_pic_pr, qn("p:nvPr"))

    # Blip fill
    blip_fill = etree.SubElement(pic, qn("p:blipFill"))
    blip = etree.SubElement(blip_fill, qn("a:blip"))
    blip.set(qn("r:embed"), img_r_id)

    stretch = etree.SubElement(blip_fill, qn("a:stretch"))
    etree.SubElement(stretch, qn("a:fillRect"))

    # Shape properties
    sp_pr = etree.SubElement(pic, qn("p:spPr"))

    xfrm = etree.SubElement(sp_pr, qn("a:xfrm"))
    off = etree.SubElement(xfrm, qn("a:off"))
    off.set("x", str(x))
    off.set("y", str(y))

    ext_elem = etree.SubElement(xfrm, qn("a:ext"))
    ext_elem.set("cx", str(cx))
    ext_elem.set("cy", str(cy))

    prst_geom = etree.SubElement(sp_pr, qn("a:prstGeom"))
    prst_geom.set("prst", "rect")
    etree.SubElement(prst_geom, qn("a:avLst"))

    # Save master
    xml_bytes = etree.tostring(
        master_elem,
        xml_declaration=True,
        encoding="UTF-8",
        standalone=True,
    )

    master_content_type = "application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml"
    pkg.set_part(master_path, xml_bytes, master_content_type)
    presentation._dirty = True

    return True
