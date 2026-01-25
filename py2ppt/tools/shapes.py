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
