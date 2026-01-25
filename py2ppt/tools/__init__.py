"""Tool-calling API for AI agents.

This module provides the main API for py2ppt, designed for AI/LLM
tool-calling interfaces. Each function is a discrete operation that
can be called independently.
"""

from .animation import (
    add_animation,
    get_available_animations,
    get_available_transitions,
    remove_animations,
    remove_transition,
    set_slide_transition,
)
from .content import (
    add_bullet,
    add_text_box,
    set_body,
    set_placeholder_text,
    set_subtitle,
    set_title,
)
from .diagrams import (
    add_diagram,
    get_diagram_types,
)
from .inspection import (
    describe_slide,
    get_placeholders,
    get_slide_count,
    get_theme_colors,
    get_theme_fonts,
    list_layouts,
)
from .media import (
    add_chart,
    add_image,
    add_table,
    crop_image,
    flip_image,
    merge_table_cells,
    rotate_image,
    style_table_cell,
    update_chart_data,
    update_table_cell,
)
from .presentation import (
    create_presentation,
    open_presentation,
    save_presentation,
)
from .shapes import (
    add_connector,
    add_shape,
    delete_shape,
    group_shapes,
    move_shape,
    resize_shape,
)
from .slides import (
    add_slide,
    delete_slide,
    duplicate_slide,
    reorder_slides,
)
from .style import (
    set_text_style,
)
from .theme import (
    apply_theme_colors,
    get_theme_info,
    set_theme_color,
    set_theme_font,
)

__all__ = [
    # Presentation tools
    "create_presentation",
    "open_presentation",
    "save_presentation",
    # Slide tools
    "add_slide",
    "delete_slide",
    "duplicate_slide",
    "reorder_slides",
    # Content tools
    "set_title",
    "set_subtitle",
    "set_body",
    "add_bullet",
    "set_placeholder_text",
    "add_text_box",
    # Media tools
    "add_table",
    "update_table_cell",
    "style_table_cell",
    "merge_table_cells",
    "add_image",
    "crop_image",
    "rotate_image",
    "flip_image",
    "add_chart",
    "update_chart_data",
    # Inspection tools
    "list_layouts",
    "describe_slide",
    "get_placeholders",
    "get_theme_colors",
    "get_theme_fonts",
    "get_slide_count",
    # Style tools
    "set_text_style",
    # Shape tools
    "add_shape",
    "add_connector",
    "group_shapes",
    "delete_shape",
    "move_shape",
    "resize_shape",
    # Theme tools
    "set_theme_color",
    "set_theme_font",
    "get_theme_info",
    "apply_theme_colors",
    # Animation tools
    "set_slide_transition",
    "add_animation",
    "remove_animations",
    "remove_transition",
    "get_available_transitions",
    "get_available_animations",
    # Diagram tools
    "add_diagram",
    "get_diagram_types",
]
