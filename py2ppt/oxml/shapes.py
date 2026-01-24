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
        off, ext = self.position.to_element()
        xfrm.append(off)
        xfrm.append(ext)

        # Geometry
        if self.preset_geometry:
            prst_geom = etree.SubElement(sp_pr, qn("a:prstGeom"))
            prst_geom.set("prst", self.preset_geometry)
            etree.SubElement(prst_geom, qn("a:avLst"))

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
class Picture:
    """A picture (pic element) in a slide."""

    id: int
    name: str
    position: Position = field(default_factory=Position)
    r_embed: str = ""  # Relationship ID to image part
    placeholder: PlaceholderInfo | None = None

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
        etree.SubElement(blip_fill, qn("a:stretch"))

        # Shape properties
        sp_pr = etree.SubElement(pic, qn("p:spPr"))
        xfrm = etree.SubElement(sp_pr, qn("a:xfrm"))
        off, ext = self.position.to_element()
        xfrm.append(off)
        xfrm.append(ext)

        prst_geom = etree.SubElement(sp_pr, qn("a:prstGeom"))
        prst_geom.set("prst", "rect")
        etree.SubElement(prst_geom, qn("a:avLst"))

        return pic

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
class TableCell:
    """A cell in a table."""

    text: str = ""
    row_span: int = 1
    col_span: int = 1

    def to_element(self) -> etree._Element:
        """Create a:tc element."""
        tc = etree.Element(qn("a:tc"))

        if self.row_span > 1:
            tc.set("rowSpan", str(self.row_span))
        if self.col_span > 1:
            tc.set("gridSpan", str(self.col_span))

        # Text body
        tx_body = etree.SubElement(tc, qn("a:txBody"))
        etree.SubElement(tx_body, qn("a:bodyPr"))
        etree.SubElement(tx_body, qn("a:lstStyle"))
        p = etree.SubElement(tx_body, qn("a:p"))
        r = etree.SubElement(p, qn("a:r"))
        t = etree.SubElement(r, qn("a:t"))
        t.text = str(self.text)
        etree.SubElement(p, qn("a:endParaRPr"))

        # Cell properties
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

    @property
    def num_rows(self) -> int:
        return len(self.rows)

    @property
    def num_cols(self) -> int:
        return len(self.col_widths) if self.col_widths else (len(self.rows[0]) if self.rows else 0)

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
        tbl_pr.set("firstRow", "1")
        tbl_pr.set("bandRow", "1")

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


class ShapeTree:
    """Collection of shapes on a slide (p:spTree)."""

    def __init__(self) -> None:
        self._shapes: list[Shape | Picture | Table] = []
        self._next_id: int = 2  # ID 1 is typically used for spTree itself

    @property
    def shapes(self) -> list[Shape | Picture | Table]:
        return self._shapes

    def get_shape_by_id(self, shape_id: int) -> Shape | Picture | Table | None:
        """Find shape by ID."""
        for shape in self._shapes:
            if shape.id == shape_id:
                return shape
        return None

    def get_shape_by_name(self, name: str) -> Shape | Picture | Table | None:
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

    def add_shape(self, shape: Shape | Picture | Table) -> None:
        """Add a shape to the tree."""
        if shape.id == 0:
            shape.id = self._next_id
            self._next_id += 1
        elif shape.id >= self._next_id:
            self._next_id = shape.id + 1
        self._shapes.append(shape)

    def remove_shape(self, shape: Shape | Picture | Table) -> bool:
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
            table = Table.from_element(gf_elem)
            if table:
                tree.add_shape(table)

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
