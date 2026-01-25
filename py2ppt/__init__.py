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
from .tools.animation import (
    add_animation,
    get_available_animations,
    get_available_transitions,
    remove_animations,
    remove_transition,
    set_slide_transition,
)
from .tools.content import (
    add_bullet,
    add_hyperlink,
    add_text_box,
    append_notes,
    append_shape_text,
    find_text,
    get_header_footer_settings,
    get_hyperlinks,
    get_notes,
    get_shape_text,
    get_text_columns,
    remove_hyperlink,
    replace_all,
    replace_text,
    set_body,
    set_date_visibility,
    set_footer,
    set_notes,
    set_placeholder_text,
    set_shape_text,
    set_slide_number_visibility,
    set_subtitle,
    set_text_columns,
    set_title,
)
from .tools.diagrams import (
    add_diagram,
    get_diagram_types,
)
from .tools.export import (
    check_export_dependencies,
    export_all_slides,
    export_slide_to_image,
    export_to_pdf,
)
from .tools.inspection import (
    describe_slide,
    get_all_thumbnails,
    get_placeholders,
    get_presentation_thumbnail,
    get_slide_count,
    get_slide_thumbnail,
    get_theme_colors,
    get_theme_fonts,
    list_layouts,
    save_slide_thumbnail,
)
from .tools.media import (
    add_audio,
    add_chart,
    add_image,
    add_table,
    add_video,
    crop_image,
    flip_image,
    get_media_shapes,
    merge_table_cells,
    rotate_image,
    style_table_cell,
    update_chart_data,
    update_table_cell,
)
from .tools.presentation import (
    create_presentation,
    get_custom_properties,
    get_custom_property,
    get_document_info,
    get_document_property,
    get_slide_size,
    open_presentation,
    remove_custom_property,
    save_presentation,
    set_custom_property,
    set_document_property,
    set_slide_size,
    set_slide_size_preset,
)
from .tools.shapes import (
    add_connector,
    add_glow,
    add_reflection,
    add_shadow,
    add_shape,
    bring_forward,
    bring_to_front,
    delete_shape,
    get_shape_order,
    group_shapes,
    move_shape,
    remove_effects,
    resize_shape,
    send_backward,
    send_to_back,
    set_shape_order,
)
from .tools.slides import (
    add_comment,
    add_section,
    add_slide,
    clear_slide_background,
    delete_comment,
    delete_section,
    delete_slide,
    duplicate_slide,
    get_comments,
    get_sections,
    rename_section,
    reorder_slides,
    set_slide_background,
)
from .tools.style import (
    set_text_style,
)
from .tools.theme import (
    apply_theme_colors,
    get_theme_info,
    set_theme_color,
    set_theme_font,
)

# === Utilities ===
from .utils.colors import (
    parse_color,
)
from .utils.errors import (
    InvalidTemplateError,
    LayoutNotFoundError,
    PlaceholderNotFoundError,
    Py2PptError,
    SlideNotFoundError,
)
from .utils.units import (
    Cm,
    Emu,
    Inches,
    Mm,
    Pt,
    parse_length,
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
    "get_slide_size",
    "set_slide_size",
    "set_slide_size_preset",
    # Document properties
    "set_document_property",
    "get_document_property",
    "get_document_info",
    "set_custom_property",
    "get_custom_property",
    "get_custom_properties",
    "remove_custom_property",

    # Slide tools
    "add_slide",
    "delete_slide",
    "duplicate_slide",
    "reorder_slides",
    "set_slide_background",
    "clear_slide_background",
    # Sections
    "add_section",
    "get_sections",
    "rename_section",
    "delete_section",
    # Comments
    "add_comment",
    "get_comments",
    "delete_comment",

    # Content tools
    "set_title",
    "set_subtitle",
    "set_body",
    "add_bullet",
    "set_placeholder_text",
    "add_text_box",
    # Speaker notes
    "set_notes",
    "get_notes",
    "append_notes",
    # Find/replace
    "find_text",
    "replace_text",
    "replace_all",
    # Headers/Footers
    "set_footer",
    "set_slide_number_visibility",
    "set_date_visibility",
    "get_header_footer_settings",
    # Hyperlinks
    "add_hyperlink",
    "remove_hyperlink",
    "get_hyperlinks",
    # Text columns
    "set_text_columns",
    "get_text_columns",
    # Shape text
    "set_shape_text",
    "get_shape_text",
    "append_shape_text",

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
    # Video/Audio
    "add_video",
    "add_audio",
    "get_media_shapes",

    # Inspection tools
    "list_layouts",
    "describe_slide",
    "get_placeholders",
    "get_theme_colors",
    "get_theme_fonts",
    "get_slide_count",
    # Thumbnails
    "get_presentation_thumbnail",
    "get_slide_thumbnail",
    "save_slide_thumbnail",
    "get_all_thumbnails",

    # Style tools
    "set_text_style",

    # Shape tools
    "add_shape",
    "add_connector",
    "group_shapes",
    "delete_shape",
    "move_shape",
    "resize_shape",
    # Object ordering
    "bring_to_front",
    "send_to_back",
    "bring_forward",
    "send_backward",
    "get_shape_order",
    "set_shape_order",
    # Shape effects
    "add_shadow",
    "add_glow",
    "add_reflection",
    "remove_effects",

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

    # Export tools
    "export_to_pdf",
    "export_slide_to_image",
    "export_all_slides",
    "check_export_dependencies",

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
