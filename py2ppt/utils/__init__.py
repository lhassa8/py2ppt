"""Utility functions for py2ppt."""

from .units import (
    EMU,
    Inches,
    Cm,
    Mm,
    Pt,
    Emu,
    inches_to_emu,
    cm_to_emu,
    mm_to_emu,
    pt_to_emu,
    emu_to_inches,
    emu_to_cm,
    emu_to_pt,
    parse_length,
)
from .colors import (
    parse_color,
    rgb_to_hex,
    hex_to_rgb,
    is_valid_hex_color,
)
from .errors import (
    Py2PptError,
    LayoutNotFoundError,
    SlideNotFoundError,
    PlaceholderNotFoundError,
    InvalidTemplateError,
    ToolResult,
    success,
    error,
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
