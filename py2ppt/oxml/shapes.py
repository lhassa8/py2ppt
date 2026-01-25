"""Shape XML handling for PresentationML.

Shapes in PowerPoint include:
- sp (shape): Basic shapes, text boxes, placeholders
- pic (picture): Images
- graphicFrame: Tables, charts, diagrams
- grpSp (group shape): Groups of shapes
- cxnSp (connector shape): Connector lines
"""

from __future__ import annotations

from dataclasses import dataclass, field
from enum import Enum

from lxml import etree

from .ns import qn
from .text import Paragraph, Run, RunProperties, TextBody


class PlaceholderType(Enum):
    """Placeholder type values from ST_PlaceholderType."""

    TITLE = "title"
    BODY = "body"
    CENTERED_TITLE = "ctrTitle"
    SUBTITLE = "subTitle"
    DATE = "dt"
    FOOTER = "ftr"
    SLIDE_NUMBER = "sldNum"
    CONTENT = "obj"  # Generic content
    CHART = "chart"
    TABLE = "tbl"
    CLIP_ART = "clipArt"
    DIAGRAM = "dgm"
    MEDIA = "media"
    SLIDE_IMAGE = "sldImg"
    PICTURE = "pic"
    HEADER = "hdr"


@dataclass
class Position:
    """Shape position and size in EMUs."""

    x: int = 0  # Left position
    y: int = 0  # Top position
    cx: int = 0  # Width
    cy: int = 0  # Height

    def to_element(self) -> tuple[etree._Element, etree._Element]:
        """Create a:off and a:ext elements."""
        off = etree.Element(qn("a:off"))
        off.set("x", str(self.x))
        off.set("y", str(self.y))

        ext = etree.Element(qn("a:ext"))
        ext.set("cx", str(self.cx))
        ext.set("cy", str(self.cy))

        return off, ext

    @classmethod
    def from_elements(
        cls, off: etree._Element | None, ext: etree._Element | None
    ) -> Position:
        """Parse from a:off and a:ext elements."""
        pos = cls()
        if off is not None:
            pos.x = int(off.get("x", "0"))
            pos.y = int(off.get("y", "0"))
        if ext is not None:
            pos.cx = int(ext.get("cx", "0"))
            pos.cy = int(ext.get("cy", "0"))
        return pos


@dataclass
class PlaceholderInfo:
    """Placeholder information from nvSpPr/nvPr/ph."""

    type: str | None = None  # Placeholder type
    idx: int | None = None  # Placeholder index
    sz: str | None = None  # Size hint ("full", "half", "quarter")
    orient: str | None = None  # Orientation ("horz", "vert")
    has_custom_prompt: bool = False

    def to_element(self) -> etree._Element:
        """Create p:ph element."""
        ph = etree.Element(qn("p:ph"))
        if self.type:
            ph.set("type", self.type)
        if self.idx is not None:
            ph.set("idx", str(self.idx))
        if self.sz:
            ph.set("sz", self.sz)
        if self.orient:
            ph.set("orient", self.orient)
        if self.has_custom_prompt:
            ph.set("hasCustomPrompt", "1")
        return ph

    @classmethod
    def from_element(cls, elem: etree._Element | None) -> PlaceholderInfo | None:
        """Parse p:ph element."""
        if elem is None:
            return None

        return cls(
            type=elem.get("type"),
            idx=int(elem.get("idx")) if elem.get("idx") else None,
            sz=elem.get("sz"),
            orient=elem.get("orient"),
            has_custom_prompt=elem.get("hasCustomPrompt") == "1",
        )


@dataclass
class TextFrame:
    """Text frame within a shape (wrapper around TextBody)."""

    body: TextBody = field(default_factory=TextBody)

    @property
    def text(self) -> str:
        return self.body.text

    @text.setter
    def text(self, value: str) -> None:
        self.body.text = value

    @property
    def paragraphs(self) -> list[Paragraph]:
        return self.body.paragraphs

    def clear(self) -> None:
        """Clear all text."""
        self.body.paragraphs = []

    def add_paragraph(
        self,
        text: str = "",
        level: int = 0,
        font_size: int | None = None,
        bold: bool = False,
        color: str | None = None,
    ) -> Paragraph:
        """Add a paragraph with optional formatting."""
        props = RunProperties(
            font_size=font_size * 100 if font_size else None,
            bold=bold if bold else None,
            color=color.lstrip("#") if color else None,
        )
        run = Run(text=text, properties=props)
        from .text import ParagraphProperties

        para_props = ParagraphProperties(level=level)
        para = Paragraph(runs=[run], properties=para_props)
        self.body.paragraphs.append(para)
        return para

    def to_element(self) -> etree._Element:
        return self.body.to_element()

    @classmethod
    def from_element(cls, elem: etree._Element) -> TextFrame:
        return cls(body=TextBody.from_element(elem))


@dataclass
class Shape:
    """A shape (sp element) in a slide."""

    id: int
    name: str
    position: Position = field(default_factory=Position)
    placeholder: PlaceholderInfo | None = None
    text_frame: TextFrame | None = None
    preset_geometry: str | None = None  # e.g., "rect", "ellipse"
    fill_color: str | None = None  # hex color or theme color name
    use_theme_color: bool = False  # if True, fill_color is a theme color name
    fill_transparency: int = 0  # 0-100 percentage
    rotation: int = 0  # rotation in 60000ths of a degree
    outline_color: str | None = None  # outline color
    outline_width: int | None = None  # outline width in EMUs

    def to_element(self) -> etree._Element:
        """Create p:sp element."""
        sp = etree.Element(qn("p:sp"))

        # Non-visual properties
        nv_sp_pr = etree.SubElement(sp, qn("p:nvSpPr"))

        # cNvPr (common non-visual properties)
        c_nv_pr = etree.SubElement(nv_sp_pr, qn("p:cNvPr"))
        c_nv_pr.set("id", str(self.id))
        c_nv_pr.set("name", self.name)

        # cNvSpPr (non-visual shape properties)
        c_nv_sp_pr = etree.SubElement(nv_sp_pr, qn("p:cNvSpPr"))
        if self.placeholder:
            c_nv_sp_pr.set("txBox", "1")

        # nvPr (non-visual properties)
        nv_pr = etree.SubElement(nv_sp_pr, qn("p:nvPr"))
        if self.placeholder:
            nv_pr.append(self.placeholder.to_element())

        # spPr (shape properties)
        sp_pr = etree.SubElement(sp, qn("p:spPr"))

        # Transform
        xfrm = etree.SubElement(sp_pr, qn("a:xfrm"))
        if self.rotation:
            xfrm.set("rot", str(self.rotation))
        off, ext = self.position.to_element()
        xfrm.append(off)
        xfrm.append(ext)

        # Geometry
        if self.preset_geometry:
            prst_geom = etree.SubElement(sp_pr, qn("a:prstGeom"))
            prst_geom.set("prst", self.preset_geometry)
            etree.SubElement(prst_geom, qn("a:avLst"))

        # Fill
        if self.fill_color:
            if self.use_theme_color:
                # Theme color fill
                solid_fill = etree.SubElement(sp_pr, qn("a:solidFill"))
                scheme_clr = etree.SubElement(solid_fill, qn("a:schemeClr"))
                scheme_clr.set("val", self.fill_color)
                if self.fill_transparency > 0:
                    alpha = etree.SubElement(scheme_clr, qn("a:alpha"))
                    # Transparency is inverse of alpha (100% transparent = 0% alpha)
                    alpha.set("val", str((100 - self.fill_transparency) * 1000))
            else:
                # RGB color fill
                solid_fill = etree.SubElement(sp_pr, qn("a:solidFill"))
                srgb_clr = etree.SubElement(solid_fill, qn("a:srgbClr"))
                srgb_clr.set("val", self.fill_color.lstrip("#").upper())
                if self.fill_transparency > 0:
                    alpha = etree.SubElement(srgb_clr, qn("a:alpha"))
                    alpha.set("val", str((100 - self.fill_transparency) * 1000))

        # Outline
        if self.outline_color or self.outline_width:
            ln = etree.SubElement(sp_pr, qn("a:ln"))
            if self.outline_width:
                ln.set("w", str(self.outline_width))
            if self.outline_color:
                solid_fill = etree.SubElement(ln, qn("a:solidFill"))
                if self.outline_color.startswith("#") or len(self.outline_color) == 6:
                    srgb_clr = etree.SubElement(solid_fill, qn("a:srgbClr"))
                    srgb_clr.set("val", self.outline_color.lstrip("#").upper())
                else:
                    scheme_clr = etree.SubElement(solid_fill, qn("a:schemeClr"))
                    scheme_clr.set("val", self.outline_color)

        # Text body
        if self.text_frame:
            sp.append(self.text_frame.to_element())

        return sp

    @classmethod
    def from_element(cls, elem: etree._Element) -> Shape:
        """Parse p:sp element."""
        # Get ID and name from cNvPr
        c_nv_pr = elem.find(f".//{qn('p:cNvPr')}")
        shape_id = int(c_nv_pr.get("id", "0")) if c_nv_pr is not None else 0
        name = c_nv_pr.get("name", "") if c_nv_pr is not None else ""

        # Get placeholder info
        ph_elem = elem.find(f".//{qn('p:ph')}")
        placeholder = PlaceholderInfo.from_element(ph_elem)

        # Get position
        xfrm = elem.find(f".//{qn('a:xfrm')}")
        if xfrm is not None:
            off = xfrm.find(qn("a:off"))
            ext = xfrm.find(qn("a:ext"))
            position = Position.from_elements(off, ext)
        else:
            position = Position()

        # Get text frame
        tx_body = elem.find(qn("p:txBody"))
        text_frame = TextFrame.from_element(tx_body) if tx_body is not None else None

        # Get geometry
        prst_geom = elem.find(f".//{qn('a:prstGeom')}")
        preset_geometry = prst_geom.get("prst") if prst_geom is not None else None

        return cls(
            id=shape_id,
            name=name,
            position=position,
            placeholder=placeholder,
            text_frame=text_frame,
            preset_geometry=preset_geometry,
        )


@dataclass
class CropRect:
    """Crop rectangle for images.

    Values are percentages (0-100000 in EMUs, where 100000 = 100%).
    Positive values crop inward from each edge.
    """

    left: int = 0  # Crop from left edge (percentage * 1000)
    top: int = 0  # Crop from top edge
    right: int = 0  # Crop from right edge
    bottom: int = 0  # Crop from bottom edge


@dataclass
class PictureEffects:
    """Visual effects for images.

    Attributes:
        shadow: Shadow effect (True for default, or dict with settings)
        reflection: Reflection effect (True for default)
        glow: Glow effect (color and size)
        soft_edges: Soft edge blur radius in points
        brightness: Brightness adjustment (-100 to 100)
        contrast: Contrast adjustment (-100 to 100)
    """

    shadow: bool | dict = False
    reflection: bool = False
    glow: dict | None = None
    soft_edges: int | None = None
    brightness: int = 0
    contrast: int = 0


@dataclass
class Picture:
    """A picture (pic element) in a slide.

    Attributes:
        id: Shape ID
        name: Shape name
        position: Position and size
        r_embed: Relationship ID to image part
        placeholder: Placeholder info if in placeholder
        rotation: Rotation angle in degrees
        crop: Crop rectangle
        effects: Visual effects
        flip_h: Horizontal flip
        flip_v: Vertical flip
    """

    id: int
    name: str
    position: Position = field(default_factory=Position)
    r_embed: str = ""  # Relationship ID to image part
    placeholder: PlaceholderInfo | None = None
    rotation: int = 0  # Degrees
    crop: CropRect | None = None
    effects: PictureEffects | None = None
    flip_h: bool = False
    flip_v: bool = False

    def to_element(self) -> etree._Element:
        """Create p:pic element."""
        pic = etree.Element(qn("p:pic"))

        # Non-visual properties
        nv_pic_pr = etree.SubElement(pic, qn("p:nvPicPr"))

        c_nv_pr = etree.SubElement(nv_pic_pr, qn("p:cNvPr"))
        c_nv_pr.set("id", str(self.id))
        c_nv_pr.set("name", self.name)

        etree.SubElement(nv_pic_pr, qn("p:cNvPicPr"))

        nv_pr = etree.SubElement(nv_pic_pr, qn("p:nvPr"))
        if self.placeholder:
            nv_pr.append(self.placeholder.to_element())

        # Blip fill
        blip_fill = etree.SubElement(pic, qn("p:blipFill"))
        blip = etree.SubElement(blip_fill, qn("a:blip"))
        blip.set(qn("r:embed"), self.r_embed)

        # Add effects to blip if brightness/contrast adjusted
        if self.effects and (self.effects.brightness != 0 or self.effects.contrast != 0):
            # Brightness/contrast via color transform
            lum = etree.SubElement(blip, qn("a:lum"))
            if self.effects.brightness != 0:
                lum.set("bright", str(self.effects.brightness * 1000))
            if self.effects.contrast != 0:
                lum.set("contrast", str(self.effects.contrast * 1000))

        # Crop (srcRect element)
        if self.crop:
            src_rect = etree.SubElement(blip_fill, qn("a:srcRect"))
            if self.crop.left > 0:
                src_rect.set("l", str(self.crop.left))
            if self.crop.top > 0:
                src_rect.set("t", str(self.crop.top))
            if self.crop.right > 0:
                src_rect.set("r", str(self.crop.right))
            if self.crop.bottom > 0:
                src_rect.set("b", str(self.crop.bottom))

        stretch = etree.SubElement(blip_fill, qn("a:stretch"))
        etree.SubElement(stretch, qn("a:fillRect"))

        # Shape properties
        sp_pr = etree.SubElement(pic, qn("p:spPr"))
        xfrm = etree.SubElement(sp_pr, qn("a:xfrm"))

        # Rotation
        if self.rotation != 0:
            xfrm.set("rot", str(self.rotation * 60000))  # Convert to 1/60000ths degree

        # Flips
        if self.flip_h:
            xfrm.set("flipH", "1")
        if self.flip_v:
            xfrm.set("flipV", "1")

        off, ext = self.position.to_element()
        xfrm.append(off)
        xfrm.append(ext)

        prst_geom = etree.SubElement(sp_pr, qn("a:prstGeom"))
        prst_geom.set("prst", "rect")
        etree.SubElement(prst_geom, qn("a:avLst"))

        # Effects
        if self.effects:
            self._add_effects(sp_pr)

        return pic

    def _add_effects(self, sp_pr: etree._Element) -> None:
        """Add visual effects to shape properties."""
        if not self.effects:
            return

        effect_lst = etree.SubElement(sp_pr, qn("a:effectLst"))

        # Shadow
        if self.effects.shadow:
            outer_shdw = etree.SubElement(effect_lst, qn("a:outerShdw"))
            if isinstance(self.effects.shadow, dict):
                outer_shdw.set("blurRad", str(self.effects.shadow.get("blur", 50800)))
                outer_shdw.set("dist", str(self.effects.shadow.get("distance", 38100)))
                outer_shdw.set("dir", str(self.effects.shadow.get("angle", 45) * 60000))
            else:
                # Default shadow
                outer_shdw.set("blurRad", "50800")
                outer_shdw.set("dist", "38100")
                outer_shdw.set("dir", "2700000")  # 45 degrees

            outer_shdw.set("algn", "tl")
            outer_shdw.set("rotWithShape", "0")

            srgb = etree.SubElement(outer_shdw, qn("a:srgbClr"))
            srgb.set("val", "000000")
            alpha = etree.SubElement(srgb, qn("a:alpha"))
            alpha.set("val", "43000")

        # Reflection
        if self.effects.reflection:
            refl = etree.SubElement(effect_lst, qn("a:reflection"))
            refl.set("blurRad", "6350")
            refl.set("stA", "50000")
            refl.set("endA", "300")
            refl.set("endPos", "55000")
            refl.set("dist", "50800")
            refl.set("dir", "5400000")
            refl.set("sy", "-100000")
            refl.set("algn", "bl")
            refl.set("rotWithShape", "0")

        # Glow
        if self.effects.glow:
            glow = etree.SubElement(effect_lst, qn("a:glow"))
            glow.set("rad", str(self.effects.glow.get("radius", 63500)))
            srgb = etree.SubElement(glow, qn("a:srgbClr"))
            srgb.set("val", self.effects.glow.get("color", "FFFF00").lstrip("#").upper())

        # Soft edges
        if self.effects.soft_edges:
            soft = etree.SubElement(effect_lst, qn("a:softEdge"))
            soft.set("rad", str(self.effects.soft_edges * 12700))

    @classmethod
    def from_element(cls, elem: etree._Element) -> Picture:
        """Parse p:pic element."""
        c_nv_pr = elem.find(f".//{qn('p:cNvPr')}")
        pic_id = int(c_nv_pr.get("id", "0")) if c_nv_pr is not None else 0
        name = c_nv_pr.get("name", "") if c_nv_pr is not None else ""

        ph_elem = elem.find(f".//{qn('p:ph')}")
        placeholder = PlaceholderInfo.from_element(ph_elem)

        xfrm = elem.find(f".//{qn('a:xfrm')}")
        if xfrm is not None:
            off = xfrm.find(qn("a:off"))
            ext = xfrm.find(qn("a:ext"))
            position = Position.from_elements(off, ext)
        else:
            position = Position()

        blip = elem.find(f".//{qn('a:blip')}")
        r_embed = blip.get(qn("r:embed"), "") if blip is not None else ""

        return cls(
            id=pic_id,
            name=name,
            position=position,
            r_embed=r_embed,
            placeholder=placeholder,
        )


@dataclass
class BorderStyle:
    """Border style for table cells.

    Attributes:
        width: Border width in EMUs (12700 = 1pt)
        color: Border color as hex without # (e.g., "000000")
        style: Border style ("solid", "dash", "dot", "dashDot", "none")
    """

    width: int = 12700  # 1pt in EMUs
    color: str = "000000"
    style: str = "solid"

    def to_element(self, border_type: str) -> etree._Element:
        """Create border element (a:lnT, a:lnB, a:lnL, a:lnR).

        Args:
            border_type: One of "T" (top), "B" (bottom), "L" (left), "R" (right)
        """
        border_map = {
            "top": "lnT", "T": "lnT",
            "bottom": "lnB", "B": "lnB",
            "left": "lnL", "L": "lnL",
            "right": "lnR", "R": "lnR",
        }
        tag = border_map.get(border_type, f"ln{border_type}")

        if self.style == "none":
            ln = etree.Element(qn(f"a:{tag}"))
            etree.SubElement(ln, qn("a:noFill"))
            return ln

        ln = etree.Element(qn(f"a:{tag}"))
        ln.set("w", str(self.width))
        ln.set("cap", "flat")
        ln.set("cmpd", "sng")

        # Solid fill
        solid_fill = etree.SubElement(ln, qn("a:solidFill"))
        srgb = etree.SubElement(solid_fill, qn("a:srgbClr"))
        srgb.set("val", self.color.lstrip("#").upper())

        # Line style
        style_map = {
            "solid": "solid",
            "dash": "dash",
            "dot": "dot",
            "dashDot": "dashDot",
        }
        prstDash = etree.SubElement(ln, qn("a:prstDash"))
        prstDash.set("val", style_map.get(self.style, "solid"))

        return ln


@dataclass
class CellStyle:
    """Style properties for a table cell.

    Attributes:
        background_color: Background color as hex without #
        border_top: Top border style
        border_bottom: Bottom border style
        border_left: Left border style
        border_right: Right border style
        vertical_align: Vertical alignment ("t" top, "ctr" center, "b" bottom)
        margin_left: Left margin in EMUs
        margin_right: Right margin in EMUs
        margin_top: Top margin in EMUs
        margin_bottom: Bottom margin in EMUs
        text_direction: Text direction ("horz", "vert", "vert270")
    """

    background_color: str | None = None
    border_top: BorderStyle | None = None
    border_bottom: BorderStyle | None = None
    border_left: BorderStyle | None = None
    border_right: BorderStyle | None = None
    vertical_align: str = "ctr"  # center
    margin_left: int | None = None
    margin_right: int | None = None
    margin_top: int | None = None
    margin_bottom: int | None = None
    text_direction: str | None = None

    def to_element(self) -> etree._Element:
        """Create a:tcPr element."""
        tc_pr = etree.Element(qn("a:tcPr"))

        # Margins
        if self.margin_left is not None:
            tc_pr.set("marL", str(self.margin_left))
        if self.margin_right is not None:
            tc_pr.set("marR", str(self.margin_right))
        if self.margin_top is not None:
            tc_pr.set("marT", str(self.margin_top))
        if self.margin_bottom is not None:
            tc_pr.set("marB", str(self.margin_bottom))

        # Vertical alignment
        if self.vertical_align:
            tc_pr.set("anchor", self.vertical_align)

        # Text direction
        if self.text_direction:
            tc_pr.set("vert", self.text_direction)

        # Borders (order matters in schema)
        if self.border_left:
            tc_pr.append(self.border_left.to_element("L"))
        if self.border_right:
            tc_pr.append(self.border_right.to_element("R"))
        if self.border_top:
            tc_pr.append(self.border_top.to_element("T"))
        if self.border_bottom:
            tc_pr.append(self.border_bottom.to_element("B"))

        # Background fill
        if self.background_color:
            solid_fill = etree.SubElement(tc_pr, qn("a:solidFill"))
            srgb = etree.SubElement(solid_fill, qn("a:srgbClr"))
            srgb.set("val", self.background_color.lstrip("#").upper())

        return tc_pr

    @classmethod
    def from_element(cls, elem: etree._Element | None) -> CellStyle | None:
        """Parse a:tcPr element."""
        if elem is None:
            return None

        style = cls()

        # Margins
        style.margin_left = int(elem.get("marL")) if elem.get("marL") else None
        style.margin_right = int(elem.get("marR")) if elem.get("marR") else None
        style.margin_top = int(elem.get("marT")) if elem.get("marT") else None
        style.margin_bottom = int(elem.get("marB")) if elem.get("marB") else None

        # Vertical alignment
        style.vertical_align = elem.get("anchor", "ctr")

        # Text direction
        style.text_direction = elem.get("vert")

        # Background
        solid_fill = elem.find(qn("a:solidFill"))
        if solid_fill is not None:
            srgb = solid_fill.find(qn("a:srgbClr"))
            if srgb is not None:
                style.background_color = srgb.get("val")

        return style


@dataclass
class TableCell:
    """A cell in a table."""

    text: str = ""
    row_span: int = 1
    col_span: int = 1
    is_merge_origin: bool = True  # False for cells merged into another
    style: CellStyle | None = None
    bold: bool = False
    font_size: int | None = None  # in centipoints
    color: str | None = None  # hex without #

    def to_element(self) -> etree._Element:
        """Create a:tc element."""
        tc = etree.Element(qn("a:tc"))

        if self.row_span > 1:
            tc.set("rowSpan", str(self.row_span))
        if self.col_span > 1:
            tc.set("gridSpan", str(self.col_span))
        if not self.is_merge_origin:
            # This cell is merged into another
            tc.set("hMerge", "1") if self.col_span == 0 else None
            tc.set("vMerge", "1") if self.row_span == 0 else None

        # Text body
        tx_body = etree.SubElement(tc, qn("a:txBody"))
        etree.SubElement(tx_body, qn("a:bodyPr"))
        etree.SubElement(tx_body, qn("a:lstStyle"))
        p = etree.SubElement(tx_body, qn("a:p"))
        r = etree.SubElement(p, qn("a:r"))

        # Run properties for formatting
        if self.bold or self.font_size or self.color:
            rPr = etree.SubElement(r, qn("a:rPr"))
            rPr.set("lang", "en-US")
            if self.bold:
                rPr.set("b", "1")
            if self.font_size:
                rPr.set("sz", str(self.font_size))
            if self.color:
                solid_fill = etree.SubElement(rPr, qn("a:solidFill"))
                srgb = etree.SubElement(solid_fill, qn("a:srgbClr"))
                srgb.set("val", self.color.lstrip("#").upper())

        t = etree.SubElement(r, qn("a:t"))
        t.text = str(self.text)
        etree.SubElement(p, qn("a:endParaRPr"))

        # Cell properties
        if self.style:
            tc.append(self.style.to_element())
        else:
            etree.SubElement(tc, qn("a:tcPr"))

        return tc


@dataclass
class Table:
    """A table (graphicFrame with a:tbl) in a slide."""

    id: int
    name: str
    position: Position = field(default_factory=Position)
    rows: list[list[TableCell]] = field(default_factory=list)
    col_widths: list[int] = field(default_factory=list)
    row_heights: list[int] = field(default_factory=list)
    first_row: bool = True  # Style first row as header
    banded_rows: bool = True  # Alternate row colors
    first_col: bool = False  # Style first column
    last_col: bool = False  # Style last column
    last_row: bool = False  # Style last row

    @property
    def num_rows(self) -> int:
        return len(self.rows)

    @property
    def num_cols(self) -> int:
        return len(self.col_widths) if self.col_widths else (len(self.rows[0]) if self.rows else 0)

    def get_cell(self, row: int, col: int) -> TableCell | None:
        """Get cell at specified row and column (0-indexed)."""
        if 0 <= row < len(self.rows) and 0 <= col < len(self.rows[row]):
            return self.rows[row][col]
        return None

    def set_cell(
        self,
        row: int,
        col: int,
        value: str | None = None,
        *,
        bold: bool | None = None,
        font_size: int | None = None,
        color: str | None = None,
        background: str | None = None,
    ) -> None:
        """Set cell content and/or formatting."""
        cell = self.get_cell(row, col)
        if cell is None:
            raise ValueError(f"Cell ({row}, {col}) out of range")

        if value is not None:
            cell.text = str(value)
        if bold is not None:
            cell.bold = bold
        if font_size is not None:
            cell.font_size = font_size * 100  # Convert to centipoints
        if color is not None:
            cell.color = color.lstrip("#")
        if background is not None:
            if cell.style is None:
                cell.style = CellStyle()
            cell.style.background_color = background.lstrip("#")

    def to_element(self) -> etree._Element:
        """Create p:graphicFrame element with table."""
        gf = etree.Element(qn("p:graphicFrame"))

        # Non-visual properties
        nv_gf_pr = etree.SubElement(gf, qn("p:nvGraphicFramePr"))
        c_nv_pr = etree.SubElement(nv_gf_pr, qn("p:cNvPr"))
        c_nv_pr.set("id", str(self.id))
        c_nv_pr.set("name", self.name)
        etree.SubElement(nv_gf_pr, qn("p:cNvGraphicFramePr"))
        etree.SubElement(nv_gf_pr, qn("p:nvPr"))

        # Transform
        xfrm = etree.SubElement(gf, qn("p:xfrm"))
        off, ext = self.position.to_element()
        xfrm.append(off)
        xfrm.append(ext)

        # Graphic
        graphic = etree.SubElement(gf, qn("a:graphic"))
        graphic_data = etree.SubElement(graphic, qn("a:graphicData"))
        graphic_data.set(
            "uri", "http://schemas.openxmlformats.org/drawingml/2006/table"
        )

        # Table
        tbl = etree.SubElement(graphic_data, qn("a:tbl"))

        # Table properties
        tbl_pr = etree.SubElement(tbl, qn("a:tblPr"))
        tbl_pr.set("firstRow", "1" if self.first_row else "0")
        tbl_pr.set("bandRow", "1" if self.banded_rows else "0")
        tbl_pr.set("firstCol", "1" if self.first_col else "0")
        tbl_pr.set("lastCol", "1" if self.last_col else "0")
        tbl_pr.set("lastRow", "1" if self.last_row else "0")

        # Grid
        tbl_grid = etree.SubElement(tbl, qn("a:tblGrid"))
        for width in self.col_widths:
            gc = etree.SubElement(tbl_grid, qn("a:gridCol"))
            gc.set("w", str(width))

        # Rows
        for row_idx, row in enumerate(self.rows):
            tr = etree.SubElement(tbl, qn("a:tr"))
            height = self.row_heights[row_idx] if row_idx < len(self.row_heights) else 370840
            tr.set("h", str(height))

            for cell in row:
                tr.append(cell.to_element())

        return gf

    @classmethod
    def from_element(cls, elem: etree._Element) -> Table | None:
        """Parse p:graphicFrame element containing a table."""
        # Check if this is a table
        tbl = elem.find(f".//{qn('a:tbl')}")
        if tbl is None:
            return None

        c_nv_pr = elem.find(f".//{qn('p:cNvPr')}")
        tbl_id = int(c_nv_pr.get("id", "0")) if c_nv_pr is not None else 0
        name = c_nv_pr.get("name", "") if c_nv_pr is not None else ""

        xfrm = elem.find(f".//{qn('p:xfrm')}")
        if xfrm is not None:
            off = xfrm.find(qn("a:off"))
            ext = xfrm.find(qn("a:ext"))
            position = Position.from_elements(off, ext)
        else:
            position = Position()

        # Parse grid
        col_widths = []
        for gc in tbl.findall(f".//{qn('a:gridCol')}"):
            col_widths.append(int(gc.get("w", "0")))

        # Parse rows
        rows = []
        row_heights = []
        for tr in tbl.findall(qn("a:tr")):
            row_heights.append(int(tr.get("h", "370840")))
            row = []
            for tc in tr.findall(qn("a:tc")):
                text = ""
                t_elem = tc.find(f".//{qn('a:t')}")
                if t_elem is not None and t_elem.text:
                    text = t_elem.text
                row_span = int(tc.get("rowSpan", "1"))
                col_span = int(tc.get("gridSpan", "1"))
                row.append(TableCell(text=text, row_span=row_span, col_span=col_span))
            rows.append(row)

        return cls(
            id=tbl_id,
            name=name,
            position=position,
            rows=rows,
            col_widths=col_widths,
            row_heights=row_heights,
        )


@dataclass
class Chart:
    """A chart (graphicFrame with c:chart) in a slide."""

    id: int
    name: str
    position: Position = field(default_factory=Position)
    r_embed: str = ""  # Relationship ID to chart part
    placeholder: PlaceholderInfo | None = None

    def to_element(self) -> etree._Element:
        """Create p:graphicFrame element with chart reference."""
        gf = etree.Element(qn("p:graphicFrame"))

        # Non-visual properties
        nv_gf_pr = etree.SubElement(gf, qn("p:nvGraphicFramePr"))
        c_nv_pr = etree.SubElement(nv_gf_pr, qn("p:cNvPr"))
        c_nv_pr.set("id", str(self.id))
        c_nv_pr.set("name", self.name)

        c_nv_gf_pr = etree.SubElement(nv_gf_pr, qn("p:cNvGraphicFramePr"))
        gf_locks = etree.SubElement(c_nv_gf_pr, qn("a:graphicFrameLocks"))
        gf_locks.set("noGrp", "1")

        nv_pr = etree.SubElement(nv_gf_pr, qn("p:nvPr"))
        if self.placeholder:
            nv_pr.append(self.placeholder.to_element())

        # Transform
        xfrm = etree.SubElement(gf, qn("p:xfrm"))
        off, ext = self.position.to_element()
        xfrm.append(off)
        xfrm.append(ext)

        # Graphic with chart reference
        graphic = etree.SubElement(gf, qn("a:graphic"))
        graphic_data = etree.SubElement(graphic, qn("a:graphicData"))
        graphic_data.set(
            "uri", "http://schemas.openxmlformats.org/drawingml/2006/chart"
        )

        # Chart reference
        chart = etree.SubElement(
            graphic_data,
            qn("c:chart"),
            nsmap={"c": "http://schemas.openxmlformats.org/drawingml/2006/chart"},
        )
        chart.set(qn("r:id"), self.r_embed)

        return gf

    @classmethod
    def from_element(cls, elem: etree._Element) -> Chart | None:
        """Parse p:graphicFrame element containing a chart."""
        # Check if this is a chart (has c:chart reference)
        chart_ref = elem.find(
            f".//{qn('c:chart')}"
        )
        if chart_ref is None:
            return None

        c_nv_pr = elem.find(f".//{qn('p:cNvPr')}")
        chart_id = int(c_nv_pr.get("id", "0")) if c_nv_pr is not None else 0
        name = c_nv_pr.get("name", "") if c_nv_pr is not None else ""

        ph_elem = elem.find(f".//{qn('p:ph')}")
        placeholder = PlaceholderInfo.from_element(ph_elem)

        xfrm = elem.find(f".//{qn('p:xfrm')}")
        if xfrm is not None:
            off = xfrm.find(qn("a:off"))
            ext = xfrm.find(qn("a:ext"))
            position = Position.from_elements(off, ext)
        else:
            position = Position()

        r_embed = chart_ref.get(qn("r:id"), "")

        return cls(
            id=chart_id,
            name=name,
            position=position,
            r_embed=r_embed,
            placeholder=placeholder,
        )


class ShapeTree:
    """Collection of shapes on a slide (p:spTree)."""

    def __init__(self) -> None:
        self._shapes: list[Shape | Picture | Table | Chart] = []
        self._next_id: int = 2  # ID 1 is typically used for spTree itself

    @property
    def shapes(self) -> list[Shape | Picture | Table | Chart]:
        return self._shapes

    def get_shape_by_id(self, shape_id: int) -> Shape | Picture | Table | Chart | None:
        """Find shape by ID."""
        for shape in self._shapes:
            if shape.id == shape_id:
                return shape
        return None

    def get_shape_by_name(self, name: str) -> Shape | Picture | Table | Chart | None:
        """Find shape by name."""
        for shape in self._shapes:
            if shape.name == name:
                return shape
        return None

    def get_placeholder(
        self, ph_type: str | None = None, ph_idx: int | None = None
    ) -> Shape | None:
        """Find a placeholder shape by type and/or index."""
        for shape in self._shapes:
            if isinstance(shape, Shape) and shape.placeholder:
                if ph_type and shape.placeholder.type != ph_type:
                    continue
                if ph_idx is not None and shape.placeholder.idx != ph_idx:
                    continue
                return shape
        return None

    def get_placeholders(self) -> list[Shape]:
        """Get all placeholder shapes."""
        return [
            s for s in self._shapes if isinstance(s, Shape) and s.placeholder is not None
        ]

    def add_shape(self, shape: Shape | Picture | Table | Chart) -> None:
        """Add a shape to the tree."""
        if shape.id == 0:
            shape.id = self._next_id
            self._next_id += 1
        elif shape.id >= self._next_id:
            self._next_id = shape.id + 1
        self._shapes.append(shape)

    def remove_shape(self, shape: Shape | Picture | Table | Chart) -> bool:
        """Remove a shape from the tree."""
        if shape in self._shapes:
            self._shapes.remove(shape)
            return True
        return False

    def to_element(self) -> etree._Element:
        """Create p:spTree element."""
        sp_tree = etree.Element(qn("p:spTree"))

        # Non-visual group shape properties
        nv_grp_sp_pr = etree.SubElement(sp_tree, qn("p:nvGrpSpPr"))
        c_nv_pr = etree.SubElement(nv_grp_sp_pr, qn("p:cNvPr"))
        c_nv_pr.set("id", "1")
        c_nv_pr.set("name", "")
        etree.SubElement(nv_grp_sp_pr, qn("p:cNvGrpSpPr"))
        etree.SubElement(nv_grp_sp_pr, qn("p:nvPr"))

        # Group shape properties
        grp_sp_pr = etree.SubElement(sp_tree, qn("p:grpSpPr"))
        xfrm = etree.SubElement(grp_sp_pr, qn("a:xfrm"))
        off = etree.SubElement(xfrm, qn("a:off"))
        off.set("x", "0")
        off.set("y", "0")
        ext = etree.SubElement(xfrm, qn("a:ext"))
        ext.set("cx", "0")
        ext.set("cy", "0")
        ch_off = etree.SubElement(xfrm, qn("a:chOff"))
        ch_off.set("x", "0")
        ch_off.set("y", "0")
        ch_ext = etree.SubElement(xfrm, qn("a:chExt"))
        ch_ext.set("cx", "0")
        ch_ext.set("cy", "0")

        # Shapes
        for shape in self._shapes:
            sp_tree.append(shape.to_element())

        return sp_tree

    @classmethod
    def from_element(cls, elem: etree._Element) -> ShapeTree:
        """Parse p:spTree element."""
        tree = cls()

        # Parse shapes (sp)
        for sp_elem in elem.findall(qn("p:sp")):
            shape = Shape.from_element(sp_elem)
            tree.add_shape(shape)

        # Parse pictures (pic)
        for pic_elem in elem.findall(qn("p:pic")):
            pic = Picture.from_element(pic_elem)
            tree.add_shape(pic)

        # Parse graphic frames (graphicFrame) - tables, charts
        for gf_elem in elem.findall(qn("p:graphicFrame")):
            # Try table first
            table = Table.from_element(gf_elem)
            if table:
                tree.add_shape(table)
                continue

            # Try chart
            chart = Chart.from_element(gf_elem)
            if chart:
                tree.add_shape(chart)

        return tree


def create_title_shape(
    shape_id: int,
    text: str = "",
    position: Position | None = None,
) -> Shape:
    """Create a title placeholder shape."""
    if position is None:
        position = Position(x=457200, y=274638, cx=8229600, cy=1143000)

    ph = PlaceholderInfo(type="title")
    tf = TextFrame()
    if text:
        tf.add_paragraph(text)

    return Shape(
        id=shape_id,
        name="Title",
        position=position,
        placeholder=ph,
        text_frame=tf,
        preset_geometry="rect",
    )


def create_body_shape(
    shape_id: int,
    items: list[str] | None = None,
    levels: list[int] | None = None,
    position: Position | None = None,
) -> Shape:
    """Create a body/content placeholder shape."""
    if position is None:
        position = Position(x=457200, y=1600200, cx=8229600, cy=4525963)

    ph = PlaceholderInfo(type="body", idx=1)
    tf = TextFrame()

    if items:
        if levels is None:
            levels = [0] * len(items)
        for item, level in zip(items, levels, strict=False):
            tf.add_paragraph(item, level=level)

    return Shape(
        id=shape_id,
        name="Content Placeholder",
        position=position,
        placeholder=ph,
        text_frame=tf,
        preset_geometry="rect",
    )


def create_subtitle_shape(
    shape_id: int,
    text: str = "",
    position: Position | None = None,
) -> Shape:
    """Create a subtitle placeholder shape."""
    if position is None:
        position = Position(x=1371600, y=3886200, cx=6400800, cy=1752600)

    ph = PlaceholderInfo(type="subTitle")
    tf = TextFrame()
    if text:
        tf.add_paragraph(text)

    return Shape(
        id=shape_id,
        name="Subtitle",
        position=position,
        placeholder=ph,
        text_frame=tf,
        preset_geometry="rect",
    )


def create_text_box(
    shape_id: int,
    text: str,
    position: Position,
    font_size: int | None = None,
    font_family: str | None = None,
    bold: bool = False,
    color: str | None = None,
) -> Shape:
    """Create a text box shape (non-placeholder)."""
    tf = TextFrame()
    tf.add_paragraph(
        text, font_size=font_size, bold=bold, color=color
    )

    return Shape(
        id=shape_id,
        name=f"TextBox {shape_id}",
        position=position,
        text_frame=tf,
        preset_geometry="rect",
    )
