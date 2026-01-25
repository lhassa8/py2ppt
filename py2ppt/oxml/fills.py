"""Fill types for shapes in PresentationML.

Fills define how shapes are colored/filled:
- SolidFill: Single color
- GradientFill: Color gradient (linear or radial)
- PatternFill: Pattern fill
- PictureFill: Image fill
- NoFill: Transparent/no fill
"""

from __future__ import annotations

from dataclasses import dataclass, field
from typing import Literal

from lxml import etree

from .ns import qn


@dataclass
class SolidFill:
    """Solid color fill.

    Attributes:
        color: Hex color without # (e.g., "FF0000")
        theme_color: Theme color reference (e.g., "accent1", "dk1")
        transparency: Transparency percentage (0-100, 0=opaque)
        tint: Tint percentage (-100 to 100)
        shade: Shade percentage (-100 to 100)
    """

    color: str | None = None
    theme_color: str | None = None
    transparency: int = 0
    tint: int | None = None
    shade: int | None = None

    def to_element(self) -> etree._Element:
        """Create a:solidFill element."""
        solid_fill = etree.Element(qn("a:solidFill"))

        if self.theme_color:
            clr = etree.SubElement(solid_fill, qn("a:schemeClr"))
            clr.set("val", self.theme_color)
        elif self.color:
            clr = etree.SubElement(solid_fill, qn("a:srgbClr"))
            clr.set("val", self.color.lstrip("#").upper())
        else:
            return solid_fill

        # Apply modifiers
        if self.transparency > 0:
            alpha = etree.SubElement(clr, qn("a:alpha"))
            alpha.set("val", str((100 - self.transparency) * 1000))

        if self.tint is not None:
            tint_el = etree.SubElement(clr, qn("a:tint"))
            tint_el.set("val", str(self.tint * 1000))

        if self.shade is not None:
            shade_el = etree.SubElement(clr, qn("a:shade"))
            shade_el.set("val", str(self.shade * 1000))

        return solid_fill

    @classmethod
    def from_element(cls, elem: etree._Element | None) -> SolidFill | None:
        """Parse a:solidFill element."""
        if elem is None:
            return None

        fill = cls()

        srgb = elem.find(qn("a:srgbClr"))
        if srgb is not None:
            fill.color = srgb.get("val")
            # Check for alpha
            alpha = srgb.find(qn("a:alpha"))
            if alpha is not None:
                fill.transparency = 100 - int(alpha.get("val", "100000")) // 1000

        scheme = elem.find(qn("a:schemeClr"))
        if scheme is not None:
            fill.theme_color = scheme.get("val")
            alpha = scheme.find(qn("a:alpha"))
            if alpha is not None:
                fill.transparency = 100 - int(alpha.get("val", "100000")) // 1000

        return fill


@dataclass
class GradientStop:
    """A color stop in a gradient.

    Attributes:
        position: Position along gradient (0-100, percentage)
        color: Hex color without #
        theme_color: Theme color reference
        transparency: Transparency percentage (0-100)
    """

    position: int
    color: str | None = None
    theme_color: str | None = None
    transparency: int = 0

    def to_element(self) -> etree._Element:
        """Create a:gs element."""
        gs = etree.Element(qn("a:gs"))
        gs.set("pos", str(self.position * 1000))  # Convert to 1/1000ths

        if self.theme_color:
            clr = etree.SubElement(gs, qn("a:schemeClr"))
            clr.set("val", self.theme_color)
        elif self.color:
            clr = etree.SubElement(gs, qn("a:srgbClr"))
            clr.set("val", self.color.lstrip("#").upper())
        else:
            return gs

        if self.transparency > 0:
            alpha = etree.SubElement(clr, qn("a:alpha"))
            alpha.set("val", str((100 - self.transparency) * 1000))

        return gs


@dataclass
class GradientFill:
    """Gradient color fill.

    Attributes:
        stops: List of gradient color stops
        direction: Angle in degrees (0=left-to-right, 90=top-to-bottom)
        gradient_type: "linear", "radial", "rectangular", "path"
        rotate_with_shape: Whether gradient rotates with shape
    """

    stops: list[GradientStop] = field(default_factory=list)
    direction: int = 0
    gradient_type: Literal["linear", "radial", "rectangular", "path"] = "linear"
    rotate_with_shape: bool = True

    def to_element(self) -> etree._Element:
        """Create a:gradFill element."""
        grad_fill = etree.Element(qn("a:gradFill"))

        if self.rotate_with_shape:
            grad_fill.set("rotWithShape", "1")

        # Gradient stop list
        gs_lst = etree.SubElement(grad_fill, qn("a:gsLst"))
        for stop in self.stops:
            gs_lst.append(stop.to_element())

        # Gradient path/direction
        if self.gradient_type == "linear":
            lin = etree.SubElement(grad_fill, qn("a:lin"))
            lin.set("ang", str(self.direction * 60000))  # Convert to 1/60000ths of degree
            lin.set("scaled", "1")
        elif self.gradient_type == "radial":
            path = etree.SubElement(grad_fill, qn("a:path"))
            path.set("path", "circle")
            fill_rect = etree.SubElement(path, qn("a:fillToRect"))
            fill_rect.set("l", "50000")
            fill_rect.set("t", "50000")
            fill_rect.set("r", "50000")
            fill_rect.set("b", "50000")
        elif self.gradient_type == "rectangular":
            path = etree.SubElement(grad_fill, qn("a:path"))
            path.set("path", "rect")
            fill_rect = etree.SubElement(path, qn("a:fillToRect"))
            fill_rect.set("l", "50000")
            fill_rect.set("t", "50000")
            fill_rect.set("r", "50000")
            fill_rect.set("b", "50000")

        return grad_fill

    @classmethod
    def from_element(cls, elem: etree._Element | None) -> GradientFill | None:
        """Parse a:gradFill element."""
        if elem is None:
            return None

        fill = cls()
        fill.rotate_with_shape = elem.get("rotWithShape") == "1"

        # Parse stops
        gs_lst = elem.find(qn("a:gsLst"))
        if gs_lst is not None:
            for gs in gs_lst.findall(qn("a:gs")):
                pos = int(gs.get("pos", "0")) // 1000
                stop = GradientStop(position=pos)

                srgb = gs.find(qn("a:srgbClr"))
                if srgb is not None:
                    stop.color = srgb.get("val")

                scheme = gs.find(qn("a:schemeClr"))
                if scheme is not None:
                    stop.theme_color = scheme.get("val")

                fill.stops.append(stop)

        # Parse direction
        lin = elem.find(qn("a:lin"))
        if lin is not None:
            fill.gradient_type = "linear"
            fill.direction = int(lin.get("ang", "0")) // 60000

        path = elem.find(qn("a:path"))
        if path is not None:
            path_type = path.get("path", "")
            if path_type == "circle":
                fill.gradient_type = "radial"
            elif path_type == "rect":
                fill.gradient_type = "rectangular"

        return fill


@dataclass
class NoFill:
    """No fill (transparent)."""

    def to_element(self) -> etree._Element:
        """Create a:noFill element."""
        return etree.Element(qn("a:noFill"))


@dataclass
class LineStyle:
    """Line/outline style for shapes.

    Attributes:
        width: Line width in EMUs (12700 = 1pt)
        color: Hex color without #
        theme_color: Theme color reference
        style: "solid", "dash", "dot", "dashDot", "lgDash"
        cap: "flat", "round", "square"
        join: "round", "bevel", "miter"
    """

    width: int = 12700
    color: str | None = None
    theme_color: str | None = None
    style: str = "solid"
    cap: str = "flat"
    join: str = "round"

    def to_element(self) -> etree._Element:
        """Create a:ln element."""
        ln = etree.Element(qn("a:ln"))
        ln.set("w", str(self.width))
        ln.set("cap", self.cap)

        # Fill
        if self.theme_color:
            solid_fill = etree.SubElement(ln, qn("a:solidFill"))
            scheme = etree.SubElement(solid_fill, qn("a:schemeClr"))
            scheme.set("val", self.theme_color)
        elif self.color:
            solid_fill = etree.SubElement(ln, qn("a:solidFill"))
            srgb = etree.SubElement(solid_fill, qn("a:srgbClr"))
            srgb.set("val", self.color.lstrip("#").upper())
        else:
            etree.SubElement(ln, qn("a:noFill"))
            return ln

        # Dash style
        style_map = {
            "solid": "solid",
            "dash": "dash",
            "dot": "dot",
            "dashDot": "dashDot",
            "lgDash": "lgDash",
            "lgDashDot": "lgDashDot",
            "sysDash": "sysDash",
            "sysDot": "sysDot",
        }
        prst_dash = etree.SubElement(ln, qn("a:prstDash"))
        prst_dash.set("val", style_map.get(self.style, "solid"))

        # Join
        if self.join == "round":
            etree.SubElement(ln, qn("a:round"))
        elif self.join == "bevel":
            etree.SubElement(ln, qn("a:bevel"))
        elif self.join == "miter":
            miter = etree.SubElement(ln, qn("a:miter"))
            miter.set("lim", "800000")

        return ln

    @classmethod
    def from_element(cls, elem: etree._Element | None) -> LineStyle | None:
        """Parse a:ln element."""
        if elem is None:
            return None

        style = cls()
        style.width = int(elem.get("w", "12700"))
        style.cap = elem.get("cap", "flat")

        # Color
        solid_fill = elem.find(qn("a:solidFill"))
        if solid_fill is not None:
            srgb = solid_fill.find(qn("a:srgbClr"))
            if srgb is not None:
                style.color = srgb.get("val")
            scheme = solid_fill.find(qn("a:schemeClr"))
            if scheme is not None:
                style.theme_color = scheme.get("val")

        # Dash
        prst_dash = elem.find(qn("a:prstDash"))
        if prst_dash is not None:
            style.style = prst_dash.get("val", "solid")

        # Join
        if elem.find(qn("a:round")) is not None:
            style.join = "round"
        elif elem.find(qn("a:bevel")) is not None:
            style.join = "bevel"
        elif elem.find(qn("a:miter")) is not None:
            style.join = "miter"

        return style


# Type alias for any fill
FillType = SolidFill | GradientFill | NoFill | None


def create_fill_element(fill: FillType) -> etree._Element | None:
    """Create fill element from fill object."""
    if fill is None:
        return None
    return fill.to_element()


def parse_fill_element(elem: etree._Element) -> FillType:
    """Parse fill element from parent element.

    Looks for a:solidFill, a:gradFill, or a:noFill as children.
    """
    solid = elem.find(qn("a:solidFill"))
    if solid is not None:
        return SolidFill.from_element(solid)

    grad = elem.find(qn("a:gradFill"))
    if grad is not None:
        return GradientFill.from_element(grad)

    no_fill = elem.find(qn("a:noFill"))
    if no_fill is not None:
        return NoFill()

    return None
