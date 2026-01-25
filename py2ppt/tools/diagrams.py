"""Diagram/SmartArt tool functions.

Functions for adding SmartArt-style diagrams to presentations.

Note: Due to the complexity of OOXML SmartArt format, this module
provides a simplified implementation using grouped shapes that
visually approximate common diagram types.
"""

from __future__ import annotations

from typing import Any

from ..core.presentation import Presentation
from ..oxml.shapes import Position, Shape, TextFrame
from ..oxml.slide import update_slide_in_package
from ..oxml.text import Paragraph, ParagraphProperties, Run, RunProperties, TextBody
from ..utils.units import Inches, parse_length

# Default diagram dimensions
DEFAULT_WIDTH = Inches(8)
DEFAULT_HEIGHT = Inches(4)
DEFAULT_LEFT = Inches(1)
DEFAULT_TOP = Inches(1.5)

# Color schemes for diagrams (theme colors)
DIAGRAM_COLORS = [
    "accent1",
    "accent2",
    "accent3",
    "accent4",
    "accent5",
    "accent6",
]


def add_diagram(
    presentation: Presentation,
    slide_number: int,
    diagram_type: str,
    items: list[str | dict],
    *,
    placeholder: str | None = None,
    left: str | int | float | None = None,
    top: str | int | float | None = None,
    width: str | int | float | None = None,
    height: str | int | float | None = None,
    title: str | None = None,
    color_scheme: str = "accent1",
) -> None:
    """Add a diagram to a slide.

    Creates a visual diagram from a list of items. The diagram is
    rendered using grouped shapes for maximum compatibility.

    Args:
        presentation: The presentation to modify
        slide_number: Slide number (1-indexed)
        diagram_type: Type of diagram. Options:
            - "process": Linear process flow (left to right)
            - "chevron": Chevron/arrow process
            - "cycle": Circular cycle diagram
            - "list": Vertical bulleted list
            - "pyramid": Pyramid with levels
            - "hierarchy": Organization hierarchy
            - "radial": Hub and spoke diagram
            - "venn": Overlapping circles (2-3 items)
        items: List of items for the diagram. Can be:
            - Simple strings: ["Step 1", "Step 2", "Step 3"]
            - Dicts with children (for hierarchy):
              [{"text": "CEO", "children": [{"text": "VP1"}, {"text": "VP2"}]}]
        placeholder: Optional placeholder to use for positioning
        left: Left position (default: 1 inch)
        top: Top position (default: 1.5 inches)
        width: Width (default: 8 inches)
        height: Height (default: 4 inches)
        title: Optional title above the diagram
        color_scheme: Theme color for the diagram (default: "accent1")

    Example:
        >>> # Simple process diagram
        >>> add_diagram(pres, 1, "process",
        ...     ["Research", "Design", "Build", "Test", "Launch"])

        >>> # Hierarchy diagram
        >>> add_diagram(pres, 1, "hierarchy", [
        ...     {"text": "CEO", "children": [
        ...         {"text": "CTO"},
        ...         {"text": "CFO"},
        ...         {"text": "COO"}
        ...     ]}
        ... ])

        >>> # Cycle diagram
        >>> add_diagram(pres, 1, "cycle",
        ...     ["Plan", "Do", "Check", "Act"])
    """
    valid_types = ["process", "chevron", "cycle", "list", "pyramid", "hierarchy", "radial", "venn"]
    if diagram_type.lower() not in valid_types:
        raise ValueError(
            f"Unknown diagram type: {diagram_type}. "
            f"Available: {', '.join(valid_types)}"
        )

    slide = presentation.get_slide(slide_number)
    slide_part = slide._part

    # Parse dimensions
    x = parse_length(left) if left else DEFAULT_LEFT
    y = parse_length(top) if top else DEFAULT_TOP
    cx = parse_length(width) if width else DEFAULT_WIDTH
    cy = parse_length(height) if height else DEFAULT_HEIGHT

    # If using placeholder, get its position
    if placeholder:
        placeholders = slide.get_placeholders()
        if placeholder in placeholders:
            ph = placeholders[placeholder]
            x = ph.position.x
            y = ph.position.y
            cx = ph.position.cx
            cy = ph.position.cy

    # Normalize items to list of dicts
    normalized_items = []
    for item in items:
        if isinstance(item, str):
            normalized_items.append({"text": item})
        else:
            normalized_items.append(item)

    # Generate shapes based on diagram type
    diagram_type = diagram_type.lower()
    if diagram_type == "process":
        _create_process_diagram(slide_part, normalized_items, x, y, cx, cy, color_scheme)
    elif diagram_type == "chevron":
        _create_chevron_diagram(slide_part, normalized_items, x, y, cx, cy, color_scheme)
    elif diagram_type == "cycle":
        _create_cycle_diagram(slide_part, normalized_items, x, y, cx, cy, color_scheme)
    elif diagram_type == "list":
        _create_list_diagram(slide_part, normalized_items, x, y, cx, cy, color_scheme, title)
    elif diagram_type == "pyramid":
        _create_pyramid_diagram(slide_part, normalized_items, x, y, cx, cy, color_scheme)
    elif diagram_type == "hierarchy":
        _create_hierarchy_diagram(slide_part, normalized_items, x, y, cx, cy, color_scheme)
    elif diagram_type == "radial":
        _create_radial_diagram(slide_part, normalized_items, x, y, cx, cy, color_scheme)
    elif diagram_type == "venn":
        _create_venn_diagram(slide_part, normalized_items, x, y, cx, cy, color_scheme)

    # Save changes
    update_slide_in_package(
        presentation._package,
        slide_number,
        slide_part,
    )


def get_diagram_types() -> dict[str, str]:
    """Get available diagram types with descriptions.

    Returns:
        Dict mapping type name to description

    Example:
        >>> types = get_diagram_types()
        >>> for name, desc in types.items():
        ...     print(f"{name}: {desc}")
    """
    return {
        "process": "Linear process flow (left to right arrows)",
        "chevron": "Chevron/arrow process flow",
        "cycle": "Circular cycle diagram",
        "list": "Vertical bulleted list with icons",
        "pyramid": "Pyramid with stacked levels",
        "hierarchy": "Organization/tree hierarchy",
        "radial": "Hub and spoke diagram",
        "venn": "Overlapping circles (2-3 items)",
    }


def _create_process_diagram(
    slide_part: Any,
    items: list[dict],
    x: int,
    y: int,
    cx: int,
    cy: int,
    color_scheme: str,
) -> None:
    """Create a horizontal process flow diagram."""
    n = len(items)
    if n == 0:
        return

    # Calculate box dimensions
    gap = Inches(0.2)
    arrow_width = Inches(0.3)
    total_gaps = gap * (n - 1) + arrow_width * (n - 1)
    box_width = (cx - total_gaps) // n
    box_height = cy

    shape_tree = slide_part.shape_tree

    for i, item in enumerate(items):
        # Calculate position
        box_x = x + i * (box_width + gap + arrow_width)
        box_y = y

        # Create rounded rectangle
        shape_id = shape_tree._next_id
        color_idx = i % len(DIAGRAM_COLORS)
        theme_color = DIAGRAM_COLORS[color_idx] if color_scheme == "accent1" else color_scheme

        shape = Shape(
            id=shape_id,
            name=f"Process {i + 1}",
            position=Position(x=box_x, y=box_y, cx=box_width, cy=box_height),
            text_frame=_create_text_frame(item["text"]),
            preset_geometry="roundRect",
            fill_color=theme_color,
            use_theme_color=True,
        )
        shape_tree.add_shape(shape)

        # Add arrow between boxes (except for last)
        if i < n - 1:
            arrow_x = box_x + box_width + gap // 2
            arrow_y = y + box_height // 2 - Inches(0.15)

            arrow_shape = Shape(
                id=shape_tree._next_id,
                name=f"Arrow {i + 1}",
                position=Position(
                    x=arrow_x,
                    y=arrow_y,
                    cx=arrow_width + gap,
                    cy=Inches(0.3),
                ),
                preset_geometry="rightArrow",
                fill_color="dk1",
                use_theme_color=True,
            )
            shape_tree.add_shape(arrow_shape)


def _create_chevron_diagram(
    slide_part: Any,
    items: list[dict],
    x: int,
    y: int,
    cx: int,
    cy: int,
    color_scheme: str,
) -> None:
    """Create a chevron arrow process diagram."""
    n = len(items)
    if n == 0:
        return

    # Calculate chevron dimensions
    overlap = Inches(0.3)
    chevron_width = (cx + overlap * (n - 1)) // n
    chevron_height = cy

    shape_tree = slide_part.shape_tree

    for i, item in enumerate(items):
        # Calculate position (with overlap)
        chev_x = x + i * (chevron_width - overlap)
        chev_y = y

        color_idx = i % len(DIAGRAM_COLORS)
        theme_color = DIAGRAM_COLORS[color_idx] if color_scheme == "accent1" else color_scheme

        shape = Shape(
            id=shape_tree._next_id,
            name=f"Chevron {i + 1}",
            position=Position(x=chev_x, y=chev_y, cx=chevron_width, cy=chevron_height),
            text_frame=_create_text_frame(item["text"]),
            preset_geometry="chevron",
            fill_color=theme_color,
            use_theme_color=True,
        )
        shape_tree.add_shape(shape)


def _create_cycle_diagram(
    slide_part: Any,
    items: list[dict],
    x: int,
    y: int,
    cx: int,
    cy: int,
    color_scheme: str,
) -> None:
    """Create a circular cycle diagram."""
    import math

    n = len(items)
    if n == 0:
        return

    # Calculate circle layout
    center_x = x + cx // 2
    center_y = y + cy // 2
    radius = min(cx, cy) // 2 - Inches(0.8)
    node_size = Inches(1.2)

    shape_tree = slide_part.shape_tree

    for i, item in enumerate(items):
        # Calculate position on circle
        angle = (2 * math.pi * i / n) - math.pi / 2  # Start from top
        node_x = int(center_x + radius * math.cos(angle) - node_size // 2)
        node_y = int(center_y + radius * math.sin(angle) - node_size // 2)

        color_idx = i % len(DIAGRAM_COLORS)
        theme_color = DIAGRAM_COLORS[color_idx] if color_scheme == "accent1" else color_scheme

        shape = Shape(
            id=shape_tree._next_id,
            name=f"Cycle {i + 1}",
            position=Position(x=node_x, y=node_y, cx=node_size, cy=node_size),
            text_frame=_create_text_frame(item["text"]),
            preset_geometry="ellipse",
            fill_color=theme_color,
            use_theme_color=True,
        )
        shape_tree.add_shape(shape)

        # Add curved arrow to next node
        if n > 1:
            next_i = (i + 1) % n
            next_angle = (2 * math.pi * next_i / n) - math.pi / 2
            mid_angle = (angle + next_angle) / 2
            if next_i == 0:  # Wrap around
                mid_angle = (angle + next_angle + 2 * math.pi) / 2

            arrow_radius = radius + Inches(0.3)
            arrow_x = int(center_x + arrow_radius * math.cos(mid_angle) - Inches(0.2))
            arrow_y = int(center_y + arrow_radius * math.sin(mid_angle) - Inches(0.2))

            # Rotation in 60000ths of a degree
            rotation = int((mid_angle + math.pi / 2) * 180 / math.pi * 60000)

            arrow_shape = Shape(
                id=shape_tree._next_id,
                name=f"Arrow {i + 1}",
                position=Position(
                    x=arrow_x,
                    y=arrow_y,
                    cx=Inches(0.4),
                    cy=Inches(0.3),
                ),
                preset_geometry="rightArrow",
                fill_color="dk1",
                use_theme_color=True,
                rotation=rotation,
            )
            shape_tree.add_shape(arrow_shape)


def _create_list_diagram(
    slide_part: Any,
    items: list[dict],
    x: int,
    y: int,
    cx: int,
    cy: int,
    color_scheme: str,
    title: str | None = None,
) -> None:
    """Create a vertical list diagram."""
    n = len(items)
    if n == 0:
        return

    # Calculate item dimensions
    gap = Inches(0.15)
    icon_size = Inches(0.4)
    title_height = Inches(0.5) if title else 0
    item_height = (cy - title_height - gap * (n - 1)) // n

    shape_tree = slide_part.shape_tree

    # Add title if present
    if title:
        title_shape = Shape(
            id=shape_tree._next_id,
            name="List Title",
            position=Position(x=x, y=y, cx=cx, cy=title_height),
            text_frame=_create_text_frame(title, bold=True),
            preset_geometry="rect",
        )
        shape_tree.add_shape(title_shape)
        y += title_height + gap

    for i, item in enumerate(items):
        item_y = y + i * (item_height + gap)

        # Icon (circle with number)
        color_idx = i % len(DIAGRAM_COLORS)
        theme_color = DIAGRAM_COLORS[color_idx] if color_scheme == "accent1" else color_scheme

        icon_shape = Shape(
            id=shape_tree._next_id,
            name=f"Icon {i + 1}",
            position=Position(
                x=x,
                y=item_y + (item_height - icon_size) // 2,
                cx=icon_size,
                cy=icon_size,
            ),
            text_frame=_create_text_frame(str(i + 1)),
            preset_geometry="ellipse",
            fill_color=theme_color,
            use_theme_color=True,
        )
        shape_tree.add_shape(icon_shape)

        # Text box
        text_x = x + icon_size + gap
        text_width = cx - icon_size - gap

        text_shape = Shape(
            id=shape_tree._next_id,
            name=f"Item {i + 1}",
            position=Position(x=text_x, y=item_y, cx=text_width, cy=item_height),
            text_frame=_create_text_frame(item["text"]),
            preset_geometry="rect",
        )
        shape_tree.add_shape(text_shape)


def _create_pyramid_diagram(
    slide_part: Any,
    items: list[dict],
    x: int,
    y: int,
    cx: int,
    cy: int,
    color_scheme: str,
) -> None:
    """Create a pyramid diagram."""
    n = len(items)
    if n == 0:
        return

    # Calculate level dimensions
    gap = Inches(0.1)
    level_height = (cy - gap * (n - 1)) // n

    shape_tree = slide_part.shape_tree

    for i, item in enumerate(items):
        # Each level gets narrower
        level_width = int(cx * (n - i) / n)
        level_x = x + (cx - level_width) // 2
        level_y = y + i * (level_height + gap)

        color_idx = i % len(DIAGRAM_COLORS)
        theme_color = DIAGRAM_COLORS[color_idx] if color_scheme == "accent1" else color_scheme

        shape = Shape(
            id=shape_tree._next_id,
            name=f"Level {i + 1}",
            position=Position(x=level_x, y=level_y, cx=level_width, cy=level_height),
            text_frame=_create_text_frame(item["text"]),
            preset_geometry="rect",
            fill_color=theme_color,
            use_theme_color=True,
        )
        shape_tree.add_shape(shape)


def _create_hierarchy_diagram(
    slide_part: Any,
    items: list[dict],
    x: int,
    y: int,
    cx: int,
    cy: int,
    color_scheme: str,
) -> None:
    """Create a hierarchy/org chart diagram."""
    if not items:
        return

    shape_tree = slide_part.shape_tree
    node_width = Inches(1.5)
    node_height = Inches(0.6)
    v_gap = Inches(0.4)
    h_gap = Inches(0.3)

    def add_node(item: dict, node_x: int, node_y: int, level: int) -> None:
        """Recursively add hierarchy nodes."""
        color_idx = level % len(DIAGRAM_COLORS)
        theme_color = DIAGRAM_COLORS[color_idx] if color_scheme == "accent1" else color_scheme

        # Add this node
        shape = Shape(
            id=shape_tree._next_id,
            name=f"Node {item['text'][:20]}",
            position=Position(x=node_x, y=node_y, cx=node_width, cy=node_height),
            text_frame=_create_text_frame(item["text"]),
            preset_geometry="roundRect",
            fill_color=theme_color,
            use_theme_color=True,
        )
        shape_tree.add_shape(shape)

        # Add children
        children = item.get("children", [])
        if children:
            n_children = len(children)
            total_children_width = n_children * node_width + (n_children - 1) * h_gap
            start_x = node_x + node_width // 2 - total_children_width // 2
            child_y = node_y + node_height + v_gap

            for i, child in enumerate(children):
                child_x = start_x + i * (node_width + h_gap)
                add_node(child, child_x, child_y, level + 1)

    # Start with root nodes
    n_roots = len(items)
    total_width = n_roots * node_width + (n_roots - 1) * h_gap * 2
    start_x = x + (cx - total_width) // 2

    for i, item in enumerate(items):
        root_x = start_x + i * (node_width + h_gap * 2)
        add_node(item, root_x, y, 0)


def _create_radial_diagram(
    slide_part: Any,
    items: list[dict],
    x: int,
    y: int,
    cx: int,
    cy: int,
    color_scheme: str,
) -> None:
    """Create a hub and spoke (radial) diagram."""
    import math

    n = len(items)
    if n == 0:
        return

    # First item is the center, rest are spokes
    center_x = x + cx // 2
    center_y = y + cy // 2
    center_size = Inches(1.5)
    spoke_size = Inches(1.0)
    radius = min(cx, cy) // 2 - spoke_size

    shape_tree = slide_part.shape_tree

    # Add center node
    center_shape = Shape(
        id=shape_tree._next_id,
        name="Center",
        position=Position(
            x=center_x - center_size // 2,
            y=center_y - center_size // 2,
            cx=center_size,
            cy=center_size,
        ),
        text_frame=_create_text_frame(items[0]["text"]),
        preset_geometry="ellipse",
        fill_color=color_scheme,
        use_theme_color=True,
    )
    shape_tree.add_shape(center_shape)

    # Add spoke nodes
    spoke_items = items[1:] if len(items) > 1 else items
    n_spokes = len(spoke_items)

    for i, item in enumerate(spoke_items):
        angle = (2 * math.pi * i / n_spokes) - math.pi / 2
        spoke_x = int(center_x + radius * math.cos(angle) - spoke_size // 2)
        spoke_y = int(center_y + radius * math.sin(angle) - spoke_size // 2)

        color_idx = (i + 1) % len(DIAGRAM_COLORS)
        theme_color = DIAGRAM_COLORS[color_idx] if color_scheme == "accent1" else color_scheme

        spoke_shape = Shape(
            id=shape_tree._next_id,
            name=f"Spoke {i + 1}",
            position=Position(x=spoke_x, y=spoke_y, cx=spoke_size, cy=spoke_size),
            text_frame=_create_text_frame(item["text"]),
            preset_geometry="ellipse",
            fill_color=theme_color,
            use_theme_color=True,
        )
        shape_tree.add_shape(spoke_shape)


def _create_venn_diagram(
    slide_part: Any,
    items: list[dict],
    x: int,
    y: int,
    cx: int,
    cy: int,
    color_scheme: str,
) -> None:
    """Create a Venn diagram (2-3 overlapping circles)."""
    n = min(len(items), 3)  # Max 3 circles
    if n == 0:
        return

    center_x = x + cx // 2
    center_y = y + cy // 2
    circle_size = min(cx, cy) * 2 // 3

    shape_tree = slide_part.shape_tree

    # Calculate positions for overlapping circles
    if n == 1:
        positions = [(center_x - circle_size // 2, center_y - circle_size // 2)]
    elif n == 2:
        offset = circle_size // 4
        positions = [
            (center_x - circle_size // 2 - offset, center_y - circle_size // 2),
            (center_x - circle_size // 2 + offset, center_y - circle_size // 2),
        ]
    else:  # n == 3
        offset = circle_size // 4
        positions = [
            (center_x - circle_size // 2, center_y - circle_size // 2 - offset),
            (center_x - circle_size // 2 - offset, center_y - circle_size // 2 + offset // 2),
            (center_x - circle_size // 2 + offset, center_y - circle_size // 2 + offset // 2),
        ]

    for i in range(n):
        item = items[i]
        pos_x, pos_y = positions[i]

        color_idx = i % len(DIAGRAM_COLORS)
        theme_color = DIAGRAM_COLORS[color_idx] if color_scheme == "accent1" else color_scheme

        shape = Shape(
            id=shape_tree._next_id,
            name=f"Circle {i + 1}",
            position=Position(x=pos_x, y=pos_y, cx=circle_size, cy=circle_size),
            text_frame=_create_text_frame(item["text"]),
            preset_geometry="ellipse",
            fill_color=theme_color,
            use_theme_color=True,
            fill_transparency=30,  # Semi-transparent for overlap visibility
        )
        shape_tree.add_shape(shape)


def _create_text_frame(text: str, bold: bool = False) -> TextFrame:
    """Create a simple text frame with centered text."""
    run_props = RunProperties(bold=bold)
    run = Run(text=text, properties=run_props)
    para_props = ParagraphProperties(alignment="ctr")
    paragraph = Paragraph(runs=[run], properties=para_props)
    body = TextBody(paragraphs=[paragraph])
    return TextFrame(body=body)
