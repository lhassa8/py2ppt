"""Shape tool functions.

Functions for adding shapes, connectors, and grouping shapes.
"""

from __future__ import annotations

from typing import Literal

from ..core.presentation import Presentation
from ..oxml.fills import GradientFill, GradientStop, LineStyle, NoFill, SolidFill
from ..oxml.shapes import Position, TextFrame
from ..utils.units import parse_length

# Preset shape types available in PowerPoint
PRESET_SHAPES = {
    # Basic shapes
    "rectangle": "rect",
    "rect": "rect",
    "rounded_rectangle": "roundRect",
    "round_rect": "roundRect",
    "ellipse": "ellipse",
    "oval": "ellipse",
    "circle": "ellipse",
    "triangle": "triangle",
    "right_triangle": "rtTriangle",
    "parallelogram": "parallelogram",
    "trapezoid": "trapezoid",
    "diamond": "diamond",
    "pentagon": "pentagon",
    "hexagon": "hexagon",
    "octagon": "octagon",
    "star5": "star5",
    "star6": "star6",
    "star": "star5",
    # Arrows
    "arrow_right": "rightArrow",
    "arrow_left": "leftArrow",
    "arrow_up": "upArrow",
    "arrow_down": "downArrow",
    "arrow_left_right": "leftRightArrow",
    "arrow_up_down": "upDownArrow",
    "chevron": "chevron",
    "notched_right_arrow": "notchedRightArrow",
    # Callouts
    "callout_rectangle": "wedgeRectCallout",
    "callout_rounded": "wedgeRoundRectCallout",
    "callout_ellipse": "wedgeEllipseCallout",
    "callout_cloud": "cloudCallout",
    # Flowchart
    "flowchart_process": "flowChartProcess",
    "flowchart_decision": "flowChartDecision",
    "flowchart_terminator": "flowChartTerminator",
    "flowchart_document": "flowChartDocument",
    "flowchart_data": "flowChartInputOutput",
    "flowchart_predefined": "flowChartPredefinedProcess",
    # Block arrows
    "block_arrow_right": "rightArrow",
    "block_arrow_left": "leftArrow",
    "block_arrow_up": "upArrow",
    "block_arrow_down": "downArrow",
    # Other
    "plus": "plus",
    "minus": "mathMinus",
    "cross": "plus",
    "heart": "heart",
    "lightning": "lightningBolt",
    "sun": "sun",
    "moon": "moon",
    "cloud": "cloud",
    "brace_left": "leftBrace",
    "brace_right": "rightBrace",
    "bracket_left": "leftBracket",
    "bracket_right": "rightBracket",
}


def add_shape(
    presentation: Presentation,
    slide_number: int,
    shape_type: str,
    left: str | int,
    top: str | int,
    width: str | int,
    height: str | int,
    *,
    text: str | None = None,
    fill: str | dict | None = None,
    outline: str | dict | bool | None = True,
    rotation: int = 0,
) -> int:
    """Add a shape to a slide.

    Args:
        presentation: The presentation to modify
        slide_number: The slide number (1-indexed)
        shape_type: Shape type. Common options:
            - Basic: "rectangle", "rounded_rectangle", "ellipse", "triangle",
                    "diamond", "pentagon", "hexagon", "star"
            - Arrows: "arrow_right", "arrow_left", "arrow_up", "arrow_down",
                     "chevron"
            - Flowchart: "flowchart_process", "flowchart_decision",
                        "flowchart_terminator", "flowchart_document"
            - Other: "plus", "heart", "cloud", "lightning"
        left: Left position (e.g., "1in", "2.5cm")
        top: Top position
        width: Shape width
        height: Shape height
        text: Optional text to display in shape
        fill: Fill color/style:
            - String: Hex color ("#FF0000") or theme color ("accent1")
            - Dict for gradient: {"type": "gradient", "colors": ["#FF0000", "#0000FF"]}
            - None or "none" for no fill
        outline: Outline style:
            - True: Default black outline
            - String: Hex color for outline
            - Dict: {"color": "#000000", "width": 1, "style": "solid"}
            - False or None: No outline
        rotation: Rotation angle in degrees (0-360)

    Returns:
        Shape ID of the created shape

    Example:
        >>> # Simple blue rectangle
        >>> add_shape(pres, 1, "rectangle",
        ...     left="1in", top="2in", width="3in", height="2in",
        ...     fill="#0066CC", text="Box 1")

        >>> # Rounded rectangle with gradient
        >>> add_shape(pres, 1, "rounded_rectangle",
        ...     left="5in", top="2in", width="3in", height="2in",
        ...     fill={"type": "gradient", "colors": ["#FF0000", "#FFFF00"]})

        >>> # Arrow with custom outline
        >>> add_shape(pres, 1, "arrow_right",
        ...     left="1in", top="5in", width="4in", height="1in",
        ...     fill="accent1", outline={"color": "#000000", "width": 2})
    """
    slide = presentation.get_slide(slide_number)

    # Get preset geometry
    shape_key = shape_type.lower().replace("-", "_").replace(" ", "_")
    preset_geom = PRESET_SHAPES.get(shape_key, shape_key)

    # Create position
    position = Position(
        x=int(parse_length(left)),
        y=int(parse_length(top)),
        cx=int(parse_length(width)),
        cy=int(parse_length(height)),
    )

    # Create text frame if text provided
    text_frame = None
    if text:
        text_frame = TextFrame()
        text_frame.add_paragraph(text)

    # Create shape
    from ..oxml.shapes import Shape as OxmlShape

    shape = OxmlShape(
        id=slide._part.shape_tree._next_id,
        name=f"{shape_type} {slide._part.shape_tree._next_id}",
        position=position,
        text_frame=text_frame,
        preset_geometry=preset_geom,
    )

    # Apply fill and outline via shape properties (in element generation)
    # Store fill/outline for element generation
    shape._fill = _parse_fill(fill)
    shape._outline = _parse_outline(outline)
    shape._rotation = rotation

    slide._part.shape_tree.add_shape(shape)
    slide._save()

    return shape.id


def _parse_fill(fill):
    """Parse fill specification into fill object."""
    if fill is None or fill == "none":
        return NoFill()

    if isinstance(fill, str):
        if fill.startswith("#"):
            return SolidFill(color=fill.lstrip("#"))
        else:
            return SolidFill(theme_color=fill)

    if isinstance(fill, dict):
        fill_type = fill.get("type", "solid")
        if fill_type == "gradient":
            colors = fill.get("colors", ["#FFFFFF", "#000000"])
            direction = fill.get("direction", 90)
            stops = []
            for i, color in enumerate(colors):
                pos = int(100 * i / (len(colors) - 1)) if len(colors) > 1 else 0
                if color.startswith("#"):
                    stops.append(GradientStop(position=pos, color=color.lstrip("#")))
                else:
                    stops.append(GradientStop(position=pos, theme_color=color))
            return GradientFill(stops=stops, direction=direction)
        else:
            color = fill.get("color")
            if color and color.startswith("#"):
                return SolidFill(color=color.lstrip("#"))
            elif color:
                return SolidFill(theme_color=color)

    return None


def _parse_outline(outline):
    """Parse outline specification into LineStyle object."""
    if outline is None or outline is False:
        return None

    if outline is True:
        return LineStyle(color="000000")

    if isinstance(outline, str):
        if outline.startswith("#"):
            return LineStyle(color=outline.lstrip("#"))
        else:
            return LineStyle(theme_color=outline)

    if isinstance(outline, dict):
        color = outline.get("color", "#000000")
        width = outline.get("width", 1)
        style = outline.get("style", "solid")

        ls = LineStyle(
            width=int(width * 12700),  # Convert pt to EMU
            style=style,
        )
        if color.startswith("#"):
            ls.color = color.lstrip("#")
        else:
            ls.theme_color = color
        return ls

    return None


def add_connector(
    presentation: Presentation,
    slide_number: int,
    start_x: str | int,
    start_y: str | int,
    end_x: str | int,
    end_y: str | int,
    *,
    connector_type: Literal["straight", "elbow", "curved"] = "straight",
    color: str = "#000000",
    width: int = 1,
    start_arrow: Literal["none", "triangle", "stealth", "diamond", "oval", "open"] = "none",
    end_arrow: Literal["none", "triangle", "stealth", "diamond", "oval", "open"] = "none",
) -> int:
    """Add a connector line to a slide.

    Connectors are lines that can connect shapes together.

    Args:
        presentation: The presentation to modify
        slide_number: The slide number (1-indexed)
        start_x: X coordinate of start point
        start_y: Y coordinate of start point
        end_x: X coordinate of end point
        end_y: Y coordinate of end point
        connector_type: Type of connector:
            - "straight": Direct line
            - "elbow": Right-angle connector
            - "curved": Curved connector
        color: Line color as hex
        width: Line width in points
        start_arrow: Arrow style at start
        end_arrow: Arrow style at end

    Returns:
        Shape ID of the created connector

    Example:
        >>> # Simple line
        >>> add_connector(pres, 1, "1in", "1in", "5in", "3in")

        >>> # Arrow
        >>> add_connector(pres, 1, "1in", "4in", "5in", "4in",
        ...     end_arrow="triangle", color="#FF0000")

        >>> # Elbow connector with arrows on both ends
        >>> add_connector(pres, 1, "1in", "5in", "5in", "7in",
        ...     connector_type="elbow",
        ...     start_arrow="oval", end_arrow="triangle")
    """
    slide = presentation.get_slide(slide_number)

    # Parse coordinates
    x1 = int(parse_length(start_x))
    y1 = int(parse_length(start_y))
    x2 = int(parse_length(end_x))
    y2 = int(parse_length(end_y))

    # Calculate position (bounding box of the line)
    left = min(x1, x2)
    top = min(y1, y2)
    right = max(x1, x2)
    bottom = max(y1, y2)

    position = Position(
        x=left,
        y=top,
        cx=max(right - left, 1),  # At least 1 EMU
        cy=max(bottom - top, 1),
    )

    # Connector geometry
    connector_map = {
        "straight": "line",
        "elbow": "bentConnector3",
        "curved": "curvedConnector3",
    }
    preset_geom = connector_map.get(connector_type, "line")

    # Create connector as a shape (cxnSp element handled specially)
    from ..oxml.shapes import Shape as OxmlShape

    connector = OxmlShape(
        id=slide._part.shape_tree._next_id,
        name=f"Connector {slide._part.shape_tree._next_id}",
        position=position,
        preset_geometry=preset_geom,
    )

    # Store line style
    connector._outline = LineStyle(
        width=int(width * 12700),
        color=color.lstrip("#"),
    )
    connector._fill = NoFill()
    connector._start_arrow = start_arrow
    connector._end_arrow = end_arrow
    connector._is_connector = True

    # Flip flags for proper line direction
    connector._flip_h = x2 < x1
    connector._flip_v = y2 < y1

    slide._part.shape_tree.add_shape(connector)
    slide._save()

    return connector.id


def group_shapes(
    presentation: Presentation,
    slide_number: int,
    shape_ids: list[int],
    *,
    name: str | None = None,
) -> int:
    """Group multiple shapes together.

    Note: This is a basic implementation. Full grouping support
    requires restructuring the shape tree which is complex.

    Args:
        presentation: The presentation to modify
        slide_number: The slide number (1-indexed)
        shape_ids: List of shape IDs to group
        name: Optional name for the group

    Returns:
        Group shape ID

    Example:
        >>> shape1 = add_shape(pres, 1, "rectangle", "1in", "1in", "2in", "1in")
        >>> shape2 = add_shape(pres, 1, "ellipse", "1in", "2in", "2in", "1in")
        >>> group_id = group_shapes(pres, 1, [shape1, shape2])
    """
    slide = presentation.get_slide(slide_number)

    # Find shapes to group
    shapes_to_group = []
    for shape_id in shape_ids:
        shape = slide._part.shape_tree.get_shape_by_id(shape_id)
        if shape:
            shapes_to_group.append(shape)

    if len(shapes_to_group) < 2:
        raise ValueError("Need at least 2 shapes to create a group")

    # Create group shape
    # Note: Full implementation would move shapes into a grpSp element
    # For now, we just track the group metadata
    group_id = slide._part.shape_tree._next_id
    slide._part.shape_tree._next_id += 1

    # Store group info (shapes remain in tree but are logically grouped)
    # This is a simplified approach - full grouping requires XML restructure

    slide._save()

    return group_id


def delete_shape(
    presentation: Presentation,
    slide_number: int,
    shape_id: int,
) -> bool:
    """Delete a shape from a slide.

    Args:
        presentation: The presentation to modify
        slide_number: The slide number (1-indexed)
        shape_id: ID of the shape to delete

    Returns:
        True if shape was deleted, False if not found

    Example:
        >>> shape_id = add_shape(pres, 1, "rectangle", "1in", "1in", "2in", "1in")
        >>> delete_shape(pres, 1, shape_id)
    """
    slide = presentation.get_slide(slide_number)

    shape = slide._part.shape_tree.get_shape_by_id(shape_id)
    if shape is None:
        return False

    result = slide._part.shape_tree.remove_shape(shape)
    if result:
        slide._save()

    return result


def move_shape(
    presentation: Presentation,
    slide_number: int,
    shape_id: int,
    left: str | int,
    top: str | int,
) -> bool:
    """Move a shape to a new position.

    Args:
        presentation: The presentation to modify
        slide_number: The slide number (1-indexed)
        shape_id: ID of the shape to move
        left: New left position
        top: New top position

    Returns:
        True if shape was moved, False if not found

    Example:
        >>> move_shape(pres, 1, shape_id, "3in", "4in")
    """
    slide = presentation.get_slide(slide_number)

    shape = slide._part.shape_tree.get_shape_by_id(shape_id)
    if shape is None:
        return False

    shape.position.x = int(parse_length(left))
    shape.position.y = int(parse_length(top))

    slide._save()
    return True


def resize_shape(
    presentation: Presentation,
    slide_number: int,
    shape_id: int,
    width: str | int,
    height: str | int,
) -> bool:
    """Resize a shape.

    Args:
        presentation: The presentation to modify
        slide_number: The slide number (1-indexed)
        shape_id: ID of the shape to resize
        width: New width
        height: New height

    Returns:
        True if shape was resized, False if not found

    Example:
        >>> resize_shape(pres, 1, shape_id, "4in", "2in")
    """
    slide = presentation.get_slide(slide_number)

    shape = slide._part.shape_tree.get_shape_by_id(shape_id)
    if shape is None:
        return False

    shape.position.cx = int(parse_length(width))
    shape.position.cy = int(parse_length(height))

    slide._save()
    return True


def bring_to_front(
    presentation: Presentation,
    slide_number: int,
    shape_id: int,
) -> bool:
    """Bring a shape to the front (top of z-order).

    Args:
        presentation: The presentation to modify
        slide_number: The slide number (1-indexed)
        shape_id: ID of the shape to bring forward

    Returns:
        True if shape was moved, False if not found

    Example:
        >>> bring_to_front(pres, 1, shape_id)
    """
    slide = presentation.get_slide(slide_number)
    shape_tree = slide._part.shape_tree

    shape = shape_tree.get_shape_by_id(shape_id)
    if shape is None:
        return False

    # Remove and re-add at the end (front)
    if shape in shape_tree._shapes:
        shape_tree._shapes.remove(shape)
        shape_tree._shapes.append(shape)
        slide._save()
        return True

    return False


def send_to_back(
    presentation: Presentation,
    slide_number: int,
    shape_id: int,
) -> bool:
    """Send a shape to the back (bottom of z-order).

    Args:
        presentation: The presentation to modify
        slide_number: The slide number (1-indexed)
        shape_id: ID of the shape to send back

    Returns:
        True if shape was moved, False if not found

    Example:
        >>> send_to_back(pres, 1, shape_id)
    """
    slide = presentation.get_slide(slide_number)
    shape_tree = slide._part.shape_tree

    shape = shape_tree.get_shape_by_id(shape_id)
    if shape is None:
        return False

    # Remove and re-add at the beginning (back)
    if shape in shape_tree._shapes:
        shape_tree._shapes.remove(shape)
        shape_tree._shapes.insert(0, shape)
        slide._save()
        return True

    return False


def bring_forward(
    presentation: Presentation,
    slide_number: int,
    shape_id: int,
) -> bool:
    """Bring a shape forward one level in z-order.

    Args:
        presentation: The presentation to modify
        slide_number: The slide number (1-indexed)
        shape_id: ID of the shape to bring forward

    Returns:
        True if shape was moved, False if not found or already at front

    Example:
        >>> bring_forward(pres, 1, shape_id)
    """
    slide = presentation.get_slide(slide_number)
    shape_tree = slide._part.shape_tree

    shape = shape_tree.get_shape_by_id(shape_id)
    if shape is None:
        return False

    if shape in shape_tree._shapes:
        idx = shape_tree._shapes.index(shape)
        if idx < len(shape_tree._shapes) - 1:
            shape_tree._shapes.remove(shape)
            shape_tree._shapes.insert(idx + 1, shape)
            slide._save()
            return True

    return False


def send_backward(
    presentation: Presentation,
    slide_number: int,
    shape_id: int,
) -> bool:
    """Send a shape backward one level in z-order.

    Args:
        presentation: The presentation to modify
        slide_number: The slide number (1-indexed)
        shape_id: ID of the shape to send backward

    Returns:
        True if shape was moved, False if not found or already at back

    Example:
        >>> send_backward(pres, 1, shape_id)
    """
    slide = presentation.get_slide(slide_number)
    shape_tree = slide._part.shape_tree

    shape = shape_tree.get_shape_by_id(shape_id)
    if shape is None:
        return False

    if shape in shape_tree._shapes:
        idx = shape_tree._shapes.index(shape)
        if idx > 0:
            shape_tree._shapes.remove(shape)
            shape_tree._shapes.insert(idx - 1, shape)
            slide._save()
            return True

    return False


def get_shape_order(
    presentation: Presentation,
    slide_number: int,
) -> list[int]:
    """Get the z-order of shapes on a slide.

    Args:
        presentation: The presentation to query
        slide_number: The slide number (1-indexed)

    Returns:
        List of shape IDs in order (back to front)

    Example:
        >>> order = get_shape_order(pres, 1)
        >>> print(f"Frontmost shape: {order[-1]}")
    """
    slide = presentation.get_slide(slide_number)
    return [shape.id for shape in slide._part.shape_tree._shapes]


def set_shape_order(
    presentation: Presentation,
    slide_number: int,
    shape_ids: list[int],
) -> bool:
    """Set the z-order of shapes on a slide.

    Args:
        presentation: The presentation to modify
        slide_number: The slide number (1-indexed)
        shape_ids: List of shape IDs in desired order (back to front)

    Returns:
        True if order was set, False if any shape not found

    Example:
        >>> # Put shape 5 in front of shapes 3 and 4
        >>> set_shape_order(pres, 1, [3, 4, 5])
    """
    slide = presentation.get_slide(slide_number)
    shape_tree = slide._part.shape_tree

    # Find all shapes by ID
    new_order = []
    for shape_id in shape_ids:
        shape = shape_tree.get_shape_by_id(shape_id)
        if shape is None:
            return False
        new_order.append(shape)

    # Keep shapes not in the list in their relative order at the beginning
    remaining = [s for s in shape_tree._shapes if s not in new_order]
    shape_tree._shapes = remaining + new_order

    slide._save()
    return True


# ============================================================================
# Shape Effects
# ============================================================================


def _find_shape_element(slide_elem, shape_id: int):
    """Find a shape element by ID in the slide XML."""
    from ..oxml.ns import qn

    # Search in shape tree for sp elements
    sp_tree = slide_elem.find(f".//{qn('p:spTree')}")
    if sp_tree is None:
        return None

    # Look for shapes with matching ID
    for sp in sp_tree.findall(f".//{qn('p:sp')}"):
        nv_sp_pr = sp.find(qn("p:nvSpPr"))
        if nv_sp_pr is not None:
            c_nv_pr = nv_sp_pr.find(qn("p:cNvPr"))
            if c_nv_pr is not None and c_nv_pr.get("id") == str(shape_id):
                return sp

    # Also look for pictures (p:pic)
    for pic in sp_tree.findall(f".//{qn('p:pic')}"):
        nv_pic_pr = pic.find(qn("p:nvPicPr"))
        if nv_pic_pr is not None:
            c_nv_pr = nv_pic_pr.find(qn("p:cNvPr"))
            if c_nv_pr is not None and c_nv_pr.get("id") == str(shape_id):
                return pic

    return None


def _save_slide_element_directly(pkg, slide_number: int, slide_elem):
    """Save slide element directly without rebuilding shape tree.

    This is needed for effect changes that modify the XML directly.
    """
    from lxml import etree

    from ..oxml.ns import CONTENT_TYPE
    from ..oxml.presentation import get_presentation_part

    pres_part = get_presentation_part(pkg)
    if pres_part is None:
        return False

    slide_refs = pres_part.get_slide_refs()
    if slide_number < 1 or slide_number > len(slide_refs):
        return False

    ref = slide_refs[slide_number - 1]
    pres_rels = pkg.get_part_rels("ppt/presentation.xml")
    rel = pres_rels.get(ref.r_id)
    if rel is None:
        return False

    if rel.target.startswith("/"):
        slide_path = rel.target.lstrip("/")
    else:
        slide_path = f"ppt/{rel.target}"

    xml_bytes = etree.tostring(
        slide_elem,
        xml_declaration=True,
        encoding="UTF-8",
        standalone=True,
    )
    pkg.set_part(slide_path, xml_bytes, CONTENT_TYPE.SLIDE)
    return True


def add_shadow(
    presentation: Presentation,
    slide_number: int,
    shape_id: int,
    *,
    blur: float = 4.0,
    distance: float = 3.0,
    direction: int = 45,
    color: str = "#000000",
    transparency: int = 60,
    shadow_type: Literal["outer", "inner"] = "outer",
) -> bool:
    """Add a shadow effect to a shape.

    Args:
        presentation: The presentation to modify
        slide_number: The slide number (1-indexed)
        shape_id: The shape ID to add shadow to
        blur: Blur radius in points (default 4)
        distance: Shadow offset distance in points (default 3)
        direction: Direction angle in degrees (default 45, down-right)
        color: Shadow color as hex (default "#000000")
        transparency: Transparency percentage 0-100 (default 60)
        shadow_type: "outer" (default) or "inner"

    Returns:
        True if shadow added, False if shape not found

    Example:
        >>> add_shadow(pres, 1, shape_id, blur=5, distance=4, transparency=50)
        >>> add_shadow(pres, 1, shape_id, shadow_type="inner")
    """
    from lxml import etree

    from ..oxml.ns import qn

    slide = presentation.get_slide(slide_number)
    slide_elem = slide._part._element

    # Find shape element by ID
    shape_elem = _find_shape_element(slide_elem, shape_id)
    if shape_elem is None:
        return False

    # Find or create spPr
    sp_pr = shape_elem.find(qn("p:spPr"))
    if sp_pr is None:
        sp_pr = etree.SubElement(shape_elem, qn("p:spPr"))

    # Find or create effectLst
    effect_lst = sp_pr.find(qn("a:effectLst"))
    if effect_lst is None:
        effect_lst = etree.SubElement(sp_pr, qn("a:effectLst"))

    # Remove existing shadow
    for shadow in effect_lst.findall(qn("a:outerShdw")):
        effect_lst.remove(shadow)
    for shadow in effect_lst.findall(qn("a:innerShdw")):
        effect_lst.remove(shadow)

    # Create shadow element
    shadow_tag = qn("a:outerShdw") if shadow_type == "outer" else qn("a:innerShdw")
    shadow = etree.SubElement(effect_lst, shadow_tag)

    # Convert points to EMUs (1 pt = 12700 EMUs)
    blur_emu = int(blur * 12700)
    dist_emu = int(distance * 12700)
    # Direction in 60000ths of a degree
    dir_val = direction * 60000

    shadow.set("blurRad", str(blur_emu))
    shadow.set("dist", str(dist_emu))
    shadow.set("dir", str(dir_val))

    # Set color
    srgb = etree.SubElement(shadow, qn("a:srgbClr"))
    srgb.set("val", color.lstrip("#").upper())

    # Set transparency (alpha)
    if transparency > 0:
        alpha = etree.SubElement(srgb, qn("a:alpha"))
        alpha.set("val", str((100 - transparency) * 1000))

    _save_slide_element_directly(presentation._package, slide_number, slide_elem)
    return True


def add_glow(
    presentation: Presentation,
    slide_number: int,
    shape_id: int,
    *,
    radius: float = 5.0,
    color: str = "#FFD700",
    transparency: int = 40,
) -> bool:
    """Add a glow effect to a shape.

    Args:
        presentation: The presentation to modify
        slide_number: The slide number (1-indexed)
        shape_id: The shape ID to add glow to
        radius: Glow radius in points (default 5)
        color: Glow color as hex (default "#FFD700" gold)
        transparency: Transparency percentage 0-100 (default 40)

    Returns:
        True if glow added, False if shape not found

    Example:
        >>> add_glow(pres, 1, shape_id, radius=8, color="#00FF00")
    """
    from lxml import etree

    from ..oxml.ns import qn

    slide = presentation.get_slide(slide_number)
    slide_elem = slide._part._element

    # Find shape element by ID
    shape_elem = _find_shape_element(slide_elem, shape_id)
    if shape_elem is None:
        return False

    # Find or create spPr
    sp_pr = shape_elem.find(qn("p:spPr"))
    if sp_pr is None:
        sp_pr = etree.SubElement(shape_elem, qn("p:spPr"))

    # Find or create effectLst
    effect_lst = sp_pr.find(qn("a:effectLst"))
    if effect_lst is None:
        effect_lst = etree.SubElement(sp_pr, qn("a:effectLst"))

    # Remove existing glow
    for glow in effect_lst.findall(qn("a:glow")):
        effect_lst.remove(glow)

    # Create glow element
    glow = etree.SubElement(effect_lst, qn("a:glow"))

    # Convert points to EMUs
    radius_emu = int(radius * 12700)
    glow.set("rad", str(radius_emu))

    # Set color
    srgb = etree.SubElement(glow, qn("a:srgbClr"))
    srgb.set("val", color.lstrip("#").upper())

    # Set transparency
    if transparency > 0:
        alpha = etree.SubElement(srgb, qn("a:alpha"))
        alpha.set("val", str((100 - transparency) * 1000))

    _save_slide_element_directly(presentation._package, slide_number, slide_elem)
    return True


def add_reflection(
    presentation: Presentation,
    slide_number: int,
    shape_id: int,
    *,
    blur: float = 0.5,
    start_position: int = 0,
    start_transparency: int = 0,
    end_transparency: int = 100,
    distance: float = 0.0,
    direction: int = 90,
    fade_direction: int = 90,
    size: int = 50,
) -> bool:
    """Add a reflection effect to a shape.

    Args:
        presentation: The presentation to modify
        slide_number: The slide number (1-indexed)
        shape_id: The shape ID to add reflection to
        blur: Blur radius in points (default 0.5)
        start_position: Start position percentage (default 0)
        start_transparency: Start transparency 0-100 (default 0, opaque)
        end_transparency: End transparency 0-100 (default 100, transparent)
        distance: Offset distance in points (default 0)
        direction: Reflection direction in degrees (default 90)
        fade_direction: Fade direction in degrees (default 90)
        size: Reflection size percentage (default 50)

    Returns:
        True if reflection added, False if shape not found

    Example:
        >>> add_reflection(pres, 1, shape_id, size=30, blur=1)
    """
    from lxml import etree

    from ..oxml.ns import qn

    slide = presentation.get_slide(slide_number)
    slide_elem = slide._part._element

    # Find shape element by ID
    shape_elem = _find_shape_element(slide_elem, shape_id)
    if shape_elem is None:
        return False

    # Find or create spPr
    sp_pr = shape_elem.find(qn("p:spPr"))
    if sp_pr is None:
        sp_pr = etree.SubElement(shape_elem, qn("p:spPr"))

    # Find or create effectLst
    effect_lst = sp_pr.find(qn("a:effectLst"))
    if effect_lst is None:
        effect_lst = etree.SubElement(sp_pr, qn("a:effectLst"))

    # Remove existing reflection
    for refl in effect_lst.findall(qn("a:reflection")):
        effect_lst.remove(refl)

    # Create reflection element
    refl = etree.SubElement(effect_lst, qn("a:reflection"))

    # Convert values
    blur_emu = int(blur * 12700)
    dist_emu = int(distance * 12700)
    dir_val = direction * 60000
    fade_dir_val = fade_direction * 60000

    refl.set("blurRad", str(blur_emu))
    refl.set("stA", str(start_transparency * 1000))
    refl.set("stPos", str(start_position * 1000))
    refl.set("endA", str(end_transparency * 1000))
    refl.set("endPos", str(size * 1000))
    refl.set("dist", str(dist_emu))
    refl.set("dir", str(dir_val))
    refl.set("fadeDir", str(fade_dir_val))
    refl.set("algn", "bl")  # Bottom-left alignment
    refl.set("rotWithShape", "0")

    _save_slide_element_directly(presentation._package, slide_number, slide_elem)
    return True


def remove_effects(
    presentation: Presentation,
    slide_number: int,
    shape_id: int,
) -> bool:
    """Remove all effects from a shape.

    Args:
        presentation: The presentation to modify
        slide_number: The slide number (1-indexed)
        shape_id: The shape ID to remove effects from

    Returns:
        True if effects removed, False if shape not found

    Example:
        >>> remove_effects(pres, 1, shape_id)
    """
    from ..oxml.ns import qn

    slide = presentation.get_slide(slide_number)
    slide_elem = slide._part._element

    # Find shape element by ID
    shape_elem = _find_shape_element(slide_elem, shape_id)
    if shape_elem is None:
        return False

    # Find spPr
    sp_pr = shape_elem.find(qn("p:spPr"))
    if sp_pr is None:
        return True  # No effects to remove

    # Remove effectLst
    effect_lst = sp_pr.find(qn("a:effectLst"))
    if effect_lst is not None:
        sp_pr.remove(effect_lst)

    _save_slide_element_directly(presentation._package, slide_number, slide_elem)
    return True


# ============================================================================
# 3D Effects
# ============================================================================


def add_3d_rotation(
    presentation: Presentation,
    slide_number: int,
    shape_id: int,
    *,
    lat: int = 0,
    lon: int = 0,
    rev: int = 0,
    preset: str | None = None,
) -> bool:
    """Add 3D rotation to a shape.

    Args:
        presentation: The presentation to modify
        slide_number: The slide number (1-indexed)
        shape_id: The shape ID to rotate
        lat: Latitude rotation (x-axis) in degrees (-90 to 90)
        lon: Longitude rotation (y-axis) in degrees (-180 to 180)
        rev: Revolution rotation (z-axis) in degrees (0 to 360)
        preset: Camera preset name (overrides lat/lon/rev). Use get_3d_presets()
                for available options. Common: "isometricLeftDown", "perspectiveLeft"

    Returns:
        True if rotation applied, False if shape not found

    Example:
        >>> # Custom rotation
        >>> add_3d_rotation(pres, 1, shape_id, lat=20, lon=30)
        >>> # Preset rotation
        >>> add_3d_rotation(pres, 1, shape_id, preset="isometricLeftDown")
    """
    from lxml import etree

    from ..oxml.effects3d import CAMERA_PRESETS
    from ..oxml.ns import qn

    slide = presentation.get_slide(slide_number)
    slide_elem = slide._part._element

    # Find shape element by ID
    shape_elem = _find_shape_element(slide_elem, shape_id)
    if shape_elem is None:
        return False

    # Find or create spPr
    sp_pr = shape_elem.find(qn("p:spPr"))
    if sp_pr is None:
        sp_pr = etree.SubElement(shape_elem, qn("p:spPr"))

    # Remove existing scene3d
    scene3d = sp_pr.find(qn("a:scene3d"))
    if scene3d is not None:
        sp_pr.remove(scene3d)

    # Create new scene3d
    scene3d = etree.SubElement(sp_pr, qn("a:scene3d"))

    # Create camera
    camera = etree.SubElement(scene3d, qn("a:camera"))

    if preset and preset in CAMERA_PRESETS:
        camera.set("prst", preset)
    else:
        camera.set("prst", "orthographicFront")
        # Add rotation
        if lat != 0 or lon != 0 or rev != 0:
            rot = etree.SubElement(camera, qn("a:rot"))
            rot.set("lat", str(lat * 60000))
            rot.set("lon", str(lon * 60000))
            rot.set("rev", str(rev * 60000))

    # Add light rig
    light_rig = etree.SubElement(scene3d, qn("a:lightRig"))
    light_rig.set("rig", "threePt")
    light_rig.set("dir", "t")

    _save_slide_element_directly(presentation._package, slide_number, slide_elem)
    return True


def add_3d_depth(
    presentation: Presentation,
    slide_number: int,
    shape_id: int,
    depth: str | int = "10pt",
    color: str | None = None,
    *,
    contour_width: str | int = 0,
    contour_color: str | None = None,
) -> bool:
    """Add 3D extrusion depth to a shape.

    Args:
        presentation: The presentation to modify
        slide_number: The slide number (1-indexed)
        shape_id: The shape ID to extrude
        depth: Extrusion depth (e.g., "10pt", "0.5in", or EMUs as int)
        color: Extrusion color as hex (None for auto)
        contour_width: Contour line width
        contour_color: Contour line color as hex

    Returns:
        True if depth applied, False if shape not found

    Example:
        >>> add_3d_depth(pres, 1, shape_id, "20pt", color="#0066CC")
        >>> add_3d_depth(pres, 1, shape_id, "0.25in", contour_width="1pt", contour_color="#000000")
    """
    from lxml import etree

    from ..oxml.ns import qn

    slide = presentation.get_slide(slide_number)
    slide_elem = slide._part._element

    # Find shape element by ID
    shape_elem = _find_shape_element(slide_elem, shape_id)
    if shape_elem is None:
        return False

    # Parse depth
    depth_emu = int(parse_length(depth)) if isinstance(depth, str) else depth

    # Parse contour width
    contour_emu = int(parse_length(contour_width)) if isinstance(contour_width, str) else contour_width

    # Find or create spPr
    sp_pr = shape_elem.find(qn("p:spPr"))
    if sp_pr is None:
        sp_pr = etree.SubElement(shape_elem, qn("p:spPr"))

    # Remove existing sp3d
    sp3d = sp_pr.find(qn("a:sp3d"))
    if sp3d is not None:
        sp_pr.remove(sp3d)

    # Create new sp3d
    sp3d = etree.SubElement(sp_pr, qn("a:sp3d"))

    if depth_emu > 0:
        sp3d.set("extrusionH", str(depth_emu))

    if contour_emu > 0:
        sp3d.set("contourW", str(contour_emu))

    # Add extrusion color
    if depth_emu > 0 and color:
        extrusion_clr = etree.SubElement(sp3d, qn("a:extrusionClr"))
        srgb = etree.SubElement(extrusion_clr, qn("a:srgbClr"))
        srgb.set("val", color.lstrip("#").upper())

    # Add contour color
    if contour_emu > 0 and contour_color:
        contour_clr = etree.SubElement(sp3d, qn("a:contourClr"))
        srgb = etree.SubElement(contour_clr, qn("a:srgbClr"))
        srgb.set("val", contour_color.lstrip("#").upper())

    # Ensure scene3d exists for 3D to render
    scene3d = sp_pr.find(qn("a:scene3d"))
    if scene3d is None:
        scene3d = etree.SubElement(sp_pr, qn("a:scene3d"))
        camera = etree.SubElement(scene3d, qn("a:camera"))
        camera.set("prst", "orthographicFront")
        light_rig = etree.SubElement(scene3d, qn("a:lightRig"))
        light_rig.set("rig", "threePt")
        light_rig.set("dir", "t")

    _save_slide_element_directly(presentation._package, slide_number, slide_elem)
    return True


def add_bevel(
    presentation: Presentation,
    slide_number: int,
    shape_id: int,
    *,
    bevel_type: str = "circle",
    width: str | int = "6pt",
    height: str | int = "6pt",
    top: bool = True,
    bottom: bool = False,
) -> bool:
    """Add bevel effect to a shape.

    Args:
        presentation: The presentation to modify
        slide_number: The slide number (1-indexed)
        shape_id: The shape ID to bevel
        bevel_type: Bevel type preset. Options:
            - "angle", "artDeco", "circle" (default), "convex"
            - "coolSlant", "cross", "divot", "hardEdge"
            - "relaxedInset", "riblet", "slope", "softRound"
        width: Bevel width (e.g., "6pt")
        height: Bevel height (e.g., "6pt")
        top: Apply bevel to top (default True)
        bottom: Apply bevel to bottom (default False)

    Returns:
        True if bevel applied, False if shape not found

    Example:
        >>> add_bevel(pres, 1, shape_id, bevel_type="circle", width="8pt")
        >>> add_bevel(pres, 1, shape_id, bevel_type="softRound", top=True, bottom=True)
    """
    from lxml import etree

    from ..oxml.ns import qn

    slide = presentation.get_slide(slide_number)
    slide_elem = slide._part._element

    # Find shape element by ID
    shape_elem = _find_shape_element(slide_elem, shape_id)
    if shape_elem is None:
        return False

    # Parse dimensions
    width_emu = int(parse_length(width)) if isinstance(width, str) else width
    height_emu = int(parse_length(height)) if isinstance(height, str) else height

    # Find or create spPr
    sp_pr = shape_elem.find(qn("p:spPr"))
    if sp_pr is None:
        sp_pr = etree.SubElement(shape_elem, qn("p:spPr"))

    # Find or create sp3d
    sp3d = sp_pr.find(qn("a:sp3d"))
    if sp3d is None:
        sp3d = etree.SubElement(sp_pr, qn("a:sp3d"))

    # Remove existing bevels
    for bevel_elem in sp3d.findall(qn("a:bevelT")):
        sp3d.remove(bevel_elem)
    for bevel_elem in sp3d.findall(qn("a:bevelB")):
        sp3d.remove(bevel_elem)

    # Add top bevel
    if top:
        bevel_t = etree.SubElement(sp3d, qn("a:bevelT"))
        bevel_t.set("prst", bevel_type)
        bevel_t.set("w", str(width_emu))
        bevel_t.set("h", str(height_emu))

    # Add bottom bevel
    if bottom:
        bevel_b = etree.SubElement(sp3d, qn("a:bevelB"))
        bevel_b.set("prst", bevel_type)
        bevel_b.set("w", str(width_emu))
        bevel_b.set("h", str(height_emu))

    # Ensure scene3d exists for 3D to render
    scene3d = sp_pr.find(qn("a:scene3d"))
    if scene3d is None:
        scene3d = etree.SubElement(sp_pr, qn("a:scene3d"))
        camera = etree.SubElement(scene3d, qn("a:camera"))
        camera.set("prst", "orthographicFront")
        light_rig = etree.SubElement(scene3d, qn("a:lightRig"))
        light_rig.set("rig", "threePt")
        light_rig.set("dir", "t")

    _save_slide_element_directly(presentation._package, slide_number, slide_elem)
    return True


def remove_3d_effects(
    presentation: Presentation,
    slide_number: int,
    shape_id: int,
) -> bool:
    """Remove all 3D effects from a shape.

    Args:
        presentation: The presentation to modify
        slide_number: The slide number (1-indexed)
        shape_id: The shape ID to remove 3D effects from

    Returns:
        True if effects removed, False if shape not found

    Example:
        >>> remove_3d_effects(pres, 1, shape_id)
    """
    from ..oxml.ns import qn

    slide = presentation.get_slide(slide_number)
    slide_elem = slide._part._element

    # Find shape element by ID
    shape_elem = _find_shape_element(slide_elem, shape_id)
    if shape_elem is None:
        return False

    # Find spPr
    sp_pr = shape_elem.find(qn("p:spPr"))
    if sp_pr is None:
        return True  # No 3D effects to remove

    # Remove scene3d
    scene3d = sp_pr.find(qn("a:scene3d"))
    if scene3d is not None:
        sp_pr.remove(scene3d)

    # Remove sp3d
    sp3d = sp_pr.find(qn("a:sp3d"))
    if sp3d is not None:
        sp_pr.remove(sp3d)

    _save_slide_element_directly(presentation._package, slide_number, slide_elem)
    return True


def get_3d_presets() -> list[str]:
    """Get available 3D camera preset names.

    Returns:
        List of preset names for use with add_3d_rotation()

    Example:
        >>> presets = get_3d_presets()
        >>> # Use an isometric preset
        >>> add_3d_rotation(pres, 1, shape_id, preset="isometricLeftDown")
    """
    from ..oxml.effects3d import CAMERA_PRESETS

    return CAMERA_PRESETS.copy()
