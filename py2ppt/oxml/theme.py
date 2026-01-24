"""Theme XML part handling.

The theme (ppt/theme/theme1.xml) contains:
- Color schemes (accent colors, backgrounds, text)
- Font schemes (heading and body fonts)
- Format schemes (fills, lines, effects)
"""

from __future__ import annotations

from dataclasses import dataclass

from lxml import etree

from .ns import nsmap, qn
from .package import Package


@dataclass
class ThemeColor:
    """A color in the theme."""

    name: str  # e.g., "accent1", "dk1", "lt1"
    rgb: str | None = None  # Hex RGB value
    system_color: str | None = None  # Windows system color


@dataclass
class ThemeFont:
    """A font in the theme."""

    typeface: str  # Font family name
    panose: str | None = None  # PANOSE font classification
    charset: int | None = None  # Character set


@dataclass
class FontScheme:
    """Font scheme containing major (heading) and minor (body) fonts."""

    name: str
    major_font: ThemeFont  # Heading font
    minor_font: ThemeFont  # Body font
    # Additional fonts for other scripts (Latin, EA, CS)
    major_latin: ThemeFont | None = None
    minor_latin: ThemeFont | None = None


class ThemePart:
    """Handles the theme XML part.

    The theme contains:
    - themeElements
      - clrScheme (color scheme)
      - fontScheme (font scheme)
      - fmtScheme (format scheme)
    """

    def __init__(self, element: etree._Element) -> None:
        self._element = element

    @property
    def element(self) -> etree._Element:
        return self._element

    def get_name(self) -> str:
        """Get the theme name."""
        return self._element.get("name", "")

    def get_colors(self) -> dict[str, str]:
        """Get theme colors as a dict of name -> hex RGB.

        Returns colors like:
        {
            "dk1": "000000",
            "lt1": "FFFFFF",
            "dk2": "44546A",
            "lt2": "E7E6E6",
            "accent1": "4472C4",
            "accent2": "ED7D31",
            ...
        }
        """
        colors = {}

        clr_scheme = self._element.find(f".//{qn('a:clrScheme')}")
        if clr_scheme is None:
            return colors

        # Color elements in the scheme
        color_names = [
            "dk1", "lt1", "dk2", "lt2",
            "accent1", "accent2", "accent3", "accent4",
            "accent5", "accent6", "hlink", "folHlink",
        ]

        for name in color_names:
            elem = clr_scheme.find(qn(f"a:{name}"))
            if elem is not None:
                # Look for srgbClr (specific RGB)
                srgb = elem.find(qn("a:srgbClr"))
                if srgb is not None:
                    val = srgb.get("val")
                    if val:
                        colors[name] = val
                        continue

                # Look for sysClr (system color)
                sys_clr = elem.find(qn("a:sysClr"))
                if sys_clr is not None:
                    # Use lastClr as the actual RGB
                    last_clr = sys_clr.get("lastClr")
                    if last_clr:
                        colors[name] = last_clr

        return colors

    def get_color(self, name: str) -> str | None:
        """Get a specific theme color by name.

        Args:
            name: Color name (e.g., "accent1", "dk1")

        Returns:
            Hex RGB string or None
        """
        return self.get_colors().get(name)

    def get_fonts(self) -> FontScheme:
        """Get the font scheme."""
        font_scheme = self._element.find(f".//{qn('a:fontScheme')}")
        if font_scheme is None:
            return FontScheme(
                name="Default",
                major_font=ThemeFont(typeface="Calibri Light"),
                minor_font=ThemeFont(typeface="Calibri"),
            )

        name = font_scheme.get("name", "")

        # Major font (headings)
        major = font_scheme.find(qn("a:majorFont"))
        major_latin = major.find(qn("a:latin")) if major is not None else None
        major_typeface = major_latin.get("typeface", "Calibri Light") if major_latin is not None else "Calibri Light"
        major_font = ThemeFont(typeface=major_typeface)

        # Minor font (body)
        minor = font_scheme.find(qn("a:minorFont"))
        minor_latin = minor.find(qn("a:latin")) if minor is not None else None
        minor_typeface = minor_latin.get("typeface", "Calibri") if minor_latin is not None else "Calibri"
        minor_font = ThemeFont(typeface=minor_typeface)

        return FontScheme(
            name=name,
            major_font=major_font,
            minor_font=minor_font,
        )

    def get_heading_font(self) -> str:
        """Get the heading (major) font family."""
        return self.get_fonts().major_font.typeface

    def get_body_font(self) -> str:
        """Get the body (minor) font family."""
        return self.get_fonts().minor_font.typeface

    def set_color(self, name: str, rgb: str) -> None:
        """Set a theme color.

        Args:
            name: Color name (e.g., "accent1")
            rgb: Hex RGB value (with or without #)
        """
        rgb = rgb.lstrip("#").upper()

        clr_scheme = self._element.find(f".//{qn('a:clrScheme')}")
        if clr_scheme is None:
            # Create color scheme
            theme_elements = self._element.find(qn("a:themeElements"))
            if theme_elements is None:
                theme_elements = etree.SubElement(self._element, qn("a:themeElements"))
            clr_scheme = etree.SubElement(theme_elements, qn("a:clrScheme"))
            clr_scheme.set("name", "Custom")

        # Find or create color element
        elem = clr_scheme.find(qn(f"a:{name}"))
        if elem is None:
            elem = etree.SubElement(clr_scheme, qn(f"a:{name}"))

        # Clear existing children
        for child in list(elem):
            elem.remove(child)

        # Add srgbClr
        srgb = etree.SubElement(elem, qn("a:srgbClr"))
        srgb.set("val", rgb)

    def to_xml(self) -> bytes:
        """Serialize to XML bytes."""
        return etree.tostring(
            self._element,
            xml_declaration=True,
            encoding="UTF-8",
            standalone=True,
        )

    @classmethod
    def from_xml(cls, xml_bytes: bytes) -> ThemePart:
        """Parse from XML bytes."""
        element = etree.fromstring(xml_bytes)
        return cls(element)

    @classmethod
    def new(cls, name: str = "Office Theme") -> ThemePart:
        """Create a new theme with Office defaults."""
        nsmap_theme = {
            None: nsmap["a"],
        }

        root = etree.Element(qn("a:theme"), nsmap=nsmap_theme)
        root.set("name", name)

        # Theme elements
        theme_elements = etree.SubElement(root, qn("a:themeElements"))

        # Color scheme (Office default)
        clr_scheme = etree.SubElement(theme_elements, qn("a:clrScheme"))
        clr_scheme.set("name", "Office")

        # Add default colors
        default_colors = {
            "dk1": ("sysClr", "windowText", "000000"),
            "lt1": ("sysClr", "window", "FFFFFF"),
            "dk2": ("srgbClr", None, "44546A"),
            "lt2": ("srgbClr", None, "E7E6E6"),
            "accent1": ("srgbClr", None, "4472C4"),
            "accent2": ("srgbClr", None, "ED7D31"),
            "accent3": ("srgbClr", None, "A5A5A5"),
            "accent4": ("srgbClr", None, "FFC000"),
            "accent5": ("srgbClr", None, "5B9BD5"),
            "accent6": ("srgbClr", None, "70AD47"),
            "hlink": ("srgbClr", None, "0563C1"),
            "folHlink": ("srgbClr", None, "954F72"),
        }

        for color_name, (clr_type, sys_val, rgb) in default_colors.items():
            elem = etree.SubElement(clr_scheme, qn(f"a:{color_name}"))
            if clr_type == "sysClr":
                sys_clr = etree.SubElement(elem, qn("a:sysClr"))
                sys_clr.set("val", sys_val)
                sys_clr.set("lastClr", rgb)
            else:
                srgb = etree.SubElement(elem, qn("a:srgbClr"))
                srgb.set("val", rgb)

        # Font scheme
        font_scheme = etree.SubElement(theme_elements, qn("a:fontScheme"))
        font_scheme.set("name", "Office")

        # Major font (headings)
        major = etree.SubElement(font_scheme, qn("a:majorFont"))
        latin = etree.SubElement(major, qn("a:latin"))
        latin.set("typeface", "Calibri Light")
        latin.set("panose", "020F0302020204030204")
        ea = etree.SubElement(major, qn("a:ea"))
        ea.set("typeface", "")
        cs = etree.SubElement(major, qn("a:cs"))
        cs.set("typeface", "")

        # Minor font (body)
        minor = etree.SubElement(font_scheme, qn("a:minorFont"))
        latin = etree.SubElement(minor, qn("a:latin"))
        latin.set("typeface", "Calibri")
        latin.set("panose", "020F0502020204030204")
        ea = etree.SubElement(minor, qn("a:ea"))
        ea.set("typeface", "")
        cs = etree.SubElement(minor, qn("a:cs"))
        cs.set("typeface", "")

        # Format scheme (minimal)
        fmt_scheme = etree.SubElement(theme_elements, qn("a:fmtScheme"))
        fmt_scheme.set("name", "Office")

        # Fill styles (required but minimal)
        fill_style_lst = etree.SubElement(fmt_scheme, qn("a:fillStyleLst"))
        for _ in range(3):
            solid_fill = etree.SubElement(fill_style_lst, qn("a:solidFill"))
            scheme_clr = etree.SubElement(solid_fill, qn("a:schemeClr"))
            scheme_clr.set("val", "phClr")

        # Line styles
        ln_style_lst = etree.SubElement(fmt_scheme, qn("a:lnStyleLst"))
        for _ in range(3):
            ln = etree.SubElement(ln_style_lst, qn("a:ln"))
            ln.set("w", "9525")
            solid_fill = etree.SubElement(ln, qn("a:solidFill"))
            scheme_clr = etree.SubElement(solid_fill, qn("a:schemeClr"))
            scheme_clr.set("val", "phClr")

        # Effect styles (minimal)
        effect_style_lst = etree.SubElement(fmt_scheme, qn("a:effectStyleLst"))
        for _ in range(3):
            effect_style = etree.SubElement(effect_style_lst, qn("a:effectStyle"))
            etree.SubElement(effect_style, qn("a:effectLst"))

        # Background fill styles
        bg_fill_style_lst = etree.SubElement(fmt_scheme, qn("a:bgFillStyleLst"))
        for _ in range(3):
            solid_fill = etree.SubElement(bg_fill_style_lst, qn("a:solidFill"))
            scheme_clr = etree.SubElement(solid_fill, qn("a:schemeClr"))
            scheme_clr.set("val", "phClr")

        return cls(root)


def get_theme_part(pkg: Package) -> ThemePart | None:
    """Get the theme from the package.

    Most presentations have a single theme at ppt/theme/theme1.xml.
    """
    for part_name, content in pkg.iter_parts():
        if part_name.startswith("ppt/theme/") and part_name.endswith(".xml") and "/_rels/" not in part_name:
            return ThemePart.from_xml(content)
    return None


def get_theme_colors_with_names(pkg: Package) -> dict[str, tuple[str, str]]:
    """Get theme colors with friendly names.

    Returns:
        Dict of scheme_name -> (friendly_name, hex_color)
        e.g., {"accent1": ("Blue", "#4472C4")}
    """
    theme = get_theme_part(pkg)
    if theme is None:
        return {}

    colors = theme.get_colors()

    # Friendly names for scheme colors
    friendly_names = {
        "dk1": "Dark 1",
        "lt1": "Light 1",
        "dk2": "Dark 2",
        "lt2": "Light 2",
        "accent1": "Accent 1",
        "accent2": "Accent 2",
        "accent3": "Accent 3",
        "accent4": "Accent 4",
        "accent5": "Accent 5",
        "accent6": "Accent 6",
        "hlink": "Hyperlink",
        "folHlink": "Followed Hyperlink",
    }

    result = {}
    for scheme_name, rgb in colors.items():
        friendly = friendly_names.get(scheme_name, scheme_name)
        result[scheme_name] = (friendly, f"#{rgb}")

    return result
