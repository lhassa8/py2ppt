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
    add_hyperlink,
    add_text_box,
    append_notes,
    find_text,
    get_header_footer_settings,
    get_hyperlinks,
    get_notes,
    remove_hyperlink,
    replace_all,
    replace_text,
    set_body,
    set_date_visibility,
    set_footer,
    set_notes,
    set_placeholder_text,
    set_slide_number_visibility,
    set_subtitle,
    set_title,
)
from .diagrams import (
    add_diagram,
    get_diagram_types,
)
from .export import (
    check_export_dependencies,
    export_all_slides,
    export_slide_to_image,
    export_to_pdf,
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
from .presentation import (
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
from .shapes import (
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
from .slides import (
    add_section,
    add_slide,
    clear_slide_background,
    delete_section,
    delete_slide,
    duplicate_slide,
    get_sections,
    rename_section,
    reorder_slides,
    set_slide_background,
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
]
