"""py2ppt: AI-Native PowerPoint Library for Python.

A tool-calling interface for creating and manipulating PowerPoint
presentations, designed for AI/LLM agents.

Basic Usage:
    >>> import py2ppt as ppt
    >>>
    >>> # Create a new presentation
    >>> pres = ppt.create_presentation()
    >>>
    >>> # Add slides
    >>> ppt.add_slide(pres, layout="Title Slide")
    >>> ppt.set_title(pres, 1, "Hello World")
    >>>
    >>> # Save
    >>> ppt.save_presentation(pres, "hello.pptx")

AI Agent Workflow:
    >>> import py2ppt as ppt
    >>>
    >>> # Open template and inspect
    >>> pres = ppt.create_presentation(template="corporate.pptx")
    >>> layouts = ppt.list_layouts(pres)
    >>> colors = ppt.get_theme_colors(pres)
    >>>
    >>> # Create content
    >>> ppt.add_slide(pres, layout="Title and Content")
    >>> ppt.set_title(pres, 1, "Q4 Review")
    >>> ppt.set_body(pres, 1, ["Revenue up 20%", "New markets", "High NPS"])
    >>>
    >>> # Save
    >>> ppt.save_presentation(pres, "review.pptx")
"""

__version__ = "0.1.0"

# === Core Classes ===
from .core.presentation import Presentation
from .core.slide import Slide

# === Tool Functions (Main API) ===

# Presentation tools
from .tools.presentation import (
    create_presentation,
    open_presentation,
    save_presentation,
)

# Slide tools
from .tools.slides import (
    add_slide,
    delete_slide,
    duplicate_slide,
    reorder_slides,
)

# Content tools
from .tools.content import (
    set_title,
    set_subtitle,
    set_body,
    add_bullet,
    set_placeholder_text,
    add_text_box,
)

# Media tools
from .tools.media import (
    add_table,
    update_table_cell,
    add_image,
)

# Inspection tools
from .tools.inspection import (
    list_layouts,
    describe_slide,
    get_placeholders,
    get_theme_colors,
    get_theme_fonts,
    get_slide_count,
)

# Style tools
from .tools.style import (
    set_text_style,
)

# === Utilities ===
from .utils.units import (
    Inches,
    Cm,
    Mm,
    Pt,
    Emu,
    parse_length,
)

from .utils.colors import (
    parse_color,
)

# === Errors ===
from .utils.errors import (
    Py2PptError,
    LayoutNotFoundError,
    SlideNotFoundError,
    PlaceholderNotFoundError,
    InvalidTemplateError,
)

__all__ = [
    # Version
    "__version__",

    # Core classes
    "Presentation",
    "Slide",

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

    # Units
    "Inches",
    "Cm",
    "Mm",
    "Pt",
    "Emu",
    "parse_length",

    # Colors
    "parse_color",

    # Errors
    "Py2PptError",
    "LayoutNotFoundError",
    "SlideNotFoundError",
    "PlaceholderNotFoundError",
    "InvalidTemplateError",
]
