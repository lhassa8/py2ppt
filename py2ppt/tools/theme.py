"""Theme tool functions.

Functions for modifying presentation themes, colors, and fonts.
"""

from __future__ import annotations

from lxml import etree

from ..core.presentation import Presentation
from ..oxml.ns import qn


def set_theme_color(
    presentation: Presentation,
    color_name: str,
    hex_color: str,
) -> None:
    """Set a theme color in the presentation.

    Theme colors affect all elements using that color reference.
    This allows changing the entire color scheme with a single call.

    Args:
        presentation: The presentation to modify
        color_name: Theme color name to modify. Options:
            - "dk1" / "dark1": Dark 1 (usually black text)
            - "lt1" / "light1": Light 1 (usually white background)
            - "dk2" / "dark2": Dark 2 (secondary dark)
            - "lt2" / "light2": Light 2 (secondary light)
            - "accent1" through "accent6": Accent colors
            - "hlink": Hyperlink color
            - "folHlink": Followed hyperlink color
        hex_color: New color as hex string (e.g., "#0066CC")

    Example:
        >>> # Change primary accent color
        >>> set_theme_color(pres, "accent1", "#FF5733")

        >>> # Change all accent colors for a new color scheme
        >>> set_theme_color(pres, "accent1", "#1E88E5")
        >>> set_theme_color(pres, "accent2", "#D81B60")
        >>> set_theme_color(pres, "accent3", "#004D40")
    """
    # Normalize color name
    color_map = {
        "dark1": "dk1",
        "light1": "lt1",
        "dark2": "dk2",
        "light2": "lt2",
        "hyperlink": "hlink",
        "followed_hyperlink": "folHlink",
    }
    color_name = color_map.get(color_name.lower(), color_name.lower())

    # Get theme part
    theme_part = presentation._theme
    if theme_part is None:
        raise ValueError("Presentation has no theme")

    theme_elem = theme_part._element

    # Find color scheme element
    clr_scheme = theme_elem.find(f".//{qn('a:clrScheme')}")
    if clr_scheme is None:
        raise ValueError("Theme has no color scheme")

    # Find the specific color element
    color_elem = clr_scheme.find(qn(f"a:{color_name}"))
    if color_elem is None:
        raise ValueError(f"Unknown theme color: {color_name}")

    # Clear existing color and set new one
    for child in list(color_elem):
        color_elem.remove(child)

    srgb = etree.SubElement(color_elem, qn("a:srgbClr"))
    srgb.set("val", hex_color.lstrip("#").upper())

    # Save theme
    presentation._save_theme()


def set_theme_font(
    presentation: Presentation,
    role: str,
    font_family: str,
) -> None:
    """Set a theme font in the presentation.

    Theme fonts affect all text using that font role.

    Args:
        presentation: The presentation to modify
        role: Font role to modify:
            - "major" / "heading": Font for titles and headings
            - "minor" / "body": Font for body text
        font_family: Font family name (e.g., "Calibri", "Arial", "Georgia")

    Example:
        >>> # Set heading font
        >>> set_theme_font(pres, "major", "Georgia")

        >>> # Set body font
        >>> set_theme_font(pres, "minor", "Calibri")
    """
    # Normalize role
    role_map = {
        "heading": "majorFont",
        "title": "majorFont",
        "body": "minorFont",
        "content": "minorFont",
        "major": "majorFont",
        "minor": "minorFont",
    }
    role_key = role_map.get(role.lower())
    if role_key is None:
        raise ValueError(f"Unknown font role: {role}. Use 'major' or 'minor'.")

    # Get theme part
    theme_part = presentation._theme
    if theme_part is None:
        raise ValueError("Presentation has no theme")

    theme_elem = theme_part._element

    # Find font scheme
    font_scheme = theme_elem.find(f".//{qn('a:fontScheme')}")
    if font_scheme is None:
        raise ValueError("Theme has no font scheme")

    # Find the font element (majorFont or minorFont)
    font_elem = font_scheme.find(qn(f"a:{role_key}"))
    if font_elem is None:
        raise ValueError(f"Theme has no {role_key} element")

    # Update latin font
    latin = font_elem.find(qn("a:latin"))
    if latin is not None:
        latin.set("typeface", font_family)
    else:
        latin = etree.SubElement(font_elem, qn("a:latin"))
        latin.set("typeface", font_family)

    # Save theme
    presentation._save_theme()


def get_theme_info(presentation: Presentation) -> dict:
    """Get detailed theme information.

    Returns:
        Dict with theme details including colors, fonts, and name

    Example:
        >>> info = get_theme_info(pres)
        >>> print(info["colors"]["accent1"])
        '#0066CC'
        >>> print(info["fonts"]["major"])
        'Calibri Light'
    """
    colors = presentation.get_theme_colors()
    fonts = presentation.get_theme_fonts()

    # Get theme name
    theme_name = "Default"
    theme_part = presentation._theme
    if theme_part and theme_part._element is not None:
        name = theme_part._element.get("name")
        if name:
            theme_name = name

    return {
        "name": theme_name,
        "colors": colors,
        "fonts": fonts,
    }


def apply_theme_colors(
    presentation: Presentation,
    colors: dict[str, str],
) -> None:
    """Apply multiple theme colors at once.

    This is a convenience function to set multiple colors in one call.

    Args:
        presentation: The presentation to modify
        colors: Dict mapping color names to hex values

    Example:
        >>> # Apply a custom color scheme
        >>> apply_theme_colors(pres, {
        ...     "accent1": "#1E88E5",
        ...     "accent2": "#D81B60",
        ...     "accent3": "#004D40",
        ...     "accent4": "#FFC107",
        ...     "accent5": "#7B1FA2",
        ...     "accent6": "#FF5722",
        ... })
    """
    for color_name, hex_color in colors.items():
        set_theme_color(presentation, color_name, hex_color)
