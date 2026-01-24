"""Tool-calling API for AI agents.

This module provides the main API for py2ppt, designed for AI/LLM
tool-calling interfaces. Each function is a discrete operation that
can be called independently.
"""

from .content import (
    add_bullet,
    add_text_box,
    set_body,
    set_placeholder_text,
    set_subtitle,
    set_title,
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
    add_image,
    add_table,
    update_table_cell,
)
from .presentation import (
    create_presentation,
    open_presentation,
    save_presentation,
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
    "add_image",
    # Inspection tools
    "list_layouts",
    "describe_slide",
    "get_placeholders",
    "get_theme_colors",
    "get_theme_fonts",
    "get_slide_count",
    # Style tools
    "set_text_style",
]
