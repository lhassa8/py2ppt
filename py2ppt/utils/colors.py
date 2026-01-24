"""Color parsing and conversion utilities."""

from __future__ import annotations

import re

# Named colors (subset of CSS colors)
NAMED_COLORS = {
    "black": "000000",
    "white": "FFFFFF",
    "red": "FF0000",
    "green": "00FF00",
    "blue": "0000FF",
    "yellow": "FFFF00",
    "cyan": "00FFFF",
    "magenta": "FF00FF",
    "gray": "808080",
    "grey": "808080",
    "silver": "C0C0C0",
    "maroon": "800000",
    "olive": "808000",
    "lime": "00FF00",
    "aqua": "00FFFF",
    "teal": "008080",
    "navy": "000080",
    "fuchsia": "FF00FF",
    "purple": "800080",
    "orange": "FFA500",
    "pink": "FFC0CB",
    "brown": "A52A2A",
    "gold": "FFD700",
    "coral": "FF7F50",
    "salmon": "FA8072",
    "tomato": "FF6347",
    "crimson": "DC143C",
    "indigo": "4B0082",
    "violet": "EE82EE",
    "plum": "DDA0DD",
    "orchid": "DA70D6",
    "tan": "D2B48C",
    "beige": "F5F5DC",
    "ivory": "FFFFF0",
    "khaki": "F0E68C",
    "snow": "FFFAFA",
    "azure": "F0FFFF",
    "lavender": "E6E6FA",
    "honeydew": "F0FFF0",
    "mintcream": "F5FFFA",
    "aliceblue": "F0F8FF",
    "ghostwhite": "F8F8FF",
    "seashell": "FFF5EE",
    "linen": "FAF0E6",
    "oldlace": "FDF5E6",
    "floralwhite": "FFFAF0",
    "antiquewhite": "FAEBD7",
    "papayawhip": "FFEFD5",
    "blanchedalmond": "FFEBCD",
    "bisque": "FFE4C4",
    "moccasin": "FFE4B5",
    "navajowhite": "FFDEAD",
    "peachpuff": "FFDAB9",
    "mistyrose": "FFE4E1",
    "lavenderblush": "FFF0F5",
    "cornsilk": "FFF8DC",
    "lemonchiffon": "FFFACD",
    "lightyellow": "FFFFE0",
    "lightgoldenrodyellow": "FAFAD2",
    "wheat": "F5DEB3",
}

# Hex color pattern
HEX_PATTERN = re.compile(r"^#?([0-9A-Fa-f]{3}|[0-9A-Fa-f]{6})$")

# RGB function pattern: rgb(255, 128, 64) or rgb(255 128 64)
RGB_PATTERN = re.compile(
    r"^rgb\s*\(\s*(\d{1,3})\s*[,\s]\s*(\d{1,3})\s*[,\s]\s*(\d{1,3})\s*\)$",
    re.IGNORECASE,
)


def is_valid_hex_color(color: str) -> bool:
    """Check if a string is a valid hex color.

    Args:
        color: Color string to check

    Returns:
        True if valid hex color (with or without #)
    """
    return HEX_PATTERN.match(color) is not None


def hex_to_rgb(hex_color: str) -> tuple[int, int, int]:
    """Convert hex color to RGB tuple.

    Args:
        hex_color: Hex color string (with or without #)

    Returns:
        Tuple of (R, G, B) values 0-255
    """
    hex_color = hex_color.lstrip("#")

    # Handle 3-character hex
    if len(hex_color) == 3:
        hex_color = "".join(c * 2 for c in hex_color)

    if len(hex_color) != 6:
        raise ValueError(f"Invalid hex color: {hex_color}")

    r = int(hex_color[0:2], 16)
    g = int(hex_color[2:4], 16)
    b = int(hex_color[4:6], 16)

    return (r, g, b)


def rgb_to_hex(r: int, g: int, b: int) -> str:
    """Convert RGB values to hex string (without #).

    Args:
        r: Red value 0-255
        g: Green value 0-255
        b: Blue value 0-255

    Returns:
        Hex color string without #
    """
    return f"{r:02X}{g:02X}{b:02X}"


def parse_color(color: str | tuple[int, int, int]) -> str:
    """Parse a color value to a hex string (without #).

    Supports:
    - Hex colors: "#FF6600", "FF6600", "#F60"
    - RGB tuples: (255, 102, 0)
    - RGB strings: "rgb(255, 102, 0)"
    - Named colors: "red", "blue", "orange"
    - Theme colors: "accent1", "dk1" (returned as-is with prefix)

    Args:
        color: Color value in any supported format

    Returns:
        Hex color string (uppercase, without #)

    Raises:
        ValueError: If color format is not recognized
    """
    # Handle RGB tuple
    if isinstance(color, tuple):
        if len(color) != 3:
            raise ValueError(f"RGB tuple must have 3 values, got {len(color)}")
        r, g, b = color
        return rgb_to_hex(r, g, b)

    if not isinstance(color, str):
        raise ValueError(f"Cannot parse color from {type(color).__name__}")

    color = color.strip()

    # Check for hex color
    hex_match = HEX_PATTERN.match(color)
    if hex_match:
        hex_val = hex_match.group(1)
        if len(hex_val) == 3:
            hex_val = "".join(c * 2 for c in hex_val)
        return hex_val.upper()

    # Check for rgb() function
    rgb_match = RGB_PATTERN.match(color)
    if rgb_match:
        r = int(rgb_match.group(1))
        g = int(rgb_match.group(2))
        b = int(rgb_match.group(3))
        if any(v > 255 for v in (r, g, b)):
            raise ValueError(f"RGB values must be 0-255, got ({r}, {g}, {b})")
        return rgb_to_hex(r, g, b)

    # Check for named color
    color_lower = color.lower()
    if color_lower in NAMED_COLORS:
        return NAMED_COLORS[color_lower]

    # Check for theme color names (return with prefix for later handling)
    theme_colors = [
        "dk1", "lt1", "dk2", "lt2",
        "accent1", "accent2", "accent3", "accent4",
        "accent5", "accent6", "hlink", "folHlink",
        "tx1", "tx2", "bg1", "bg2",
    ]
    if color_lower in theme_colors:
        return f"scheme:{color_lower}"

    raise ValueError(f"Unrecognized color format: {color}")


def is_theme_color(color: str) -> bool:
    """Check if a parsed color is a theme color reference.

    Args:
        color: Color string (from parse_color)

    Returns:
        True if this is a theme color reference
    """
    return color.startswith("scheme:")


def get_theme_color_name(color: str) -> str | None:
    """Get the theme color name from a parsed color.

    Args:
        color: Color string (from parse_color)

    Returns:
        Theme color name (e.g., "accent1") or None if not a theme color
    """
    if color.startswith("scheme:"):
        return color[7:]  # Remove "scheme:" prefix
    return None


def lighten(hex_color: str, amount: float = 0.2) -> str:
    """Lighten a color by a percentage.

    Args:
        hex_color: Hex color string
        amount: Amount to lighten (0.0 to 1.0)

    Returns:
        Lightened hex color
    """
    r, g, b = hex_to_rgb(hex_color)

    r = min(255, int(r + (255 - r) * amount))
    g = min(255, int(g + (255 - g) * amount))
    b = min(255, int(b + (255 - b) * amount))

    return rgb_to_hex(r, g, b)


def darken(hex_color: str, amount: float = 0.2) -> str:
    """Darken a color by a percentage.

    Args:
        hex_color: Hex color string
        amount: Amount to darken (0.0 to 1.0)

    Returns:
        Darkened hex color
    """
    r, g, b = hex_to_rgb(hex_color)

    r = max(0, int(r * (1 - amount)))
    g = max(0, int(g * (1 - amount)))
    b = max(0, int(b * (1 - amount)))

    return rgb_to_hex(r, g, b)
