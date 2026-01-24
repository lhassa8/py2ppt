"""Utility functions for py2ppt."""

from .colors import (
    hex_to_rgb,
    is_valid_hex_color,
    parse_color,
    rgb_to_hex,
)
from .errors import (
    InvalidTemplateError,
    LayoutNotFoundError,
    PlaceholderNotFoundError,
    Py2PptError,
    SlideNotFoundError,
    ToolResult,
    error,
    success,
)
from .units import (
    EMU,
    Cm,
    Emu,
    Inches,
    Mm,
    Pt,
    cm_to_emu,
    emu_to_cm,
    emu_to_inches,
    emu_to_pt,
    inches_to_emu,
    mm_to_emu,
    parse_length,
    pt_to_emu,
)

__all__ = [
    # Units
    "EMU",
    "Inches",
    "Cm",
    "Mm",
    "Pt",
    "Emu",
    "inches_to_emu",
    "cm_to_emu",
    "mm_to_emu",
    "pt_to_emu",
    "emu_to_inches",
    "emu_to_cm",
    "emu_to_pt",
    "parse_length",
    # Colors
    "parse_color",
    "rgb_to_hex",
    "hex_to_rgb",
    "is_valid_hex_color",
    # Errors
    "Py2PptError",
    "LayoutNotFoundError",
    "SlideNotFoundError",
    "PlaceholderNotFoundError",
    "InvalidTemplateError",
    "ToolResult",
    "success",
    "error",
]
