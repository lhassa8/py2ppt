"""Text and paragraph XML handling.

DrawingML text is structured as:
- txBody (text body)
  - bodyPr (body properties)
  - lstStyle (list style - optional)
  - p (paragraph) - one or more
    - pPr (paragraph properties)
    - r (run) - one or more
      - rPr (run properties - font, size, color)
      - t (text content)
"""

from __future__ import annotations

from dataclasses import dataclass, field

from lxml import etree

from .ns import qn


@dataclass
class RunProperties:
    """Run-level text formatting.

    Attributes:
        font_family: Font family name (e.g., "Arial", "Calibri")
        font_size: Font size in hundredths of a point (centipoints)
        bold: Whether text is bold
        italic: Whether text is italic
        underline: Whether text has underline
        strikethrough: Whether text has strikethrough
        double_strikethrough: Whether text has double strikethrough
        subscript: Whether text is subscript
        superscript: Whether text is superscript
        color: Hex color without # (e.g., "FF0000")
        theme_color: Theme color reference (e.g., "accent1", "dk1")
        highlight: Highlight color as hex without #
        char_spacing: Character spacing in hundredths of a point
        hyperlink: URL or relationship ID for hyperlink (starts with "rId" for internal)
        language: Language code (e.g., "en-US")
    """

    font_family: str | None = None
    font_size: int | None = None
    bold: bool | None = None
    italic: bool | None = None
    underline: bool | None = None
    strikethrough: bool | None = None
    double_strikethrough: bool | None = None
    subscript: bool | None = None
    superscript: bool | None = None
    color: str | None = None
    theme_color: str | None = None
    highlight: str | None = None
    char_spacing: int | None = None
    hyperlink: str | None = None
    language: str | None = None

    def to_element(self) -> etree._Element | None:
        """Create rPr element. Returns None if no properties set."""
        if all(
            v is None
            for v in [
                self.font_family,
                self.font_size,
                self.bold,
                self.italic,
                self.underline,
                self.strikethrough,
                self.double_strikethrough,
                self.subscript,
                self.superscript,
                self.color,
                self.theme_color,
                self.highlight,
                self.char_spacing,
                self.hyperlink,
                self.language,
            ]
        ):
            return None

        rpr = etree.Element(qn("a:rPr"))

        # Language
        if self.language is not None:
            rpr.set("lang", self.language)

        # Font size
        if self.font_size is not None:
            rpr.set("sz", str(self.font_size))

        # Bold, italic, underline
        if self.bold is not None:
            rpr.set("b", "1" if self.bold else "0")
        if self.italic is not None:
            rpr.set("i", "1" if self.italic else "0")
        if self.underline is not None:
            rpr.set("u", "sng" if self.underline else "none")

        # Strikethrough
        if self.strikethrough is not None:
            rpr.set("strike", "sngStrike" if self.strikethrough else "noStrike")
        elif self.double_strikethrough is not None:
            rpr.set("strike", "dblStrike" if self.double_strikethrough else "noStrike")

        # Superscript/Subscript (baseline adjustment)
        if self.superscript:
            rpr.set("baseline", "30000")  # 30% above baseline
        elif self.subscript:
            rpr.set("baseline", "-25000")  # 25% below baseline

        # Character spacing
        if self.char_spacing is not None:
            rpr.set("spc", str(self.char_spacing))

        # Color
        if self.color is not None:
            solid_fill = etree.SubElement(rpr, qn("a:solidFill"))
            srgb = etree.SubElement(solid_fill, qn("a:srgbClr"))
            srgb.set("val", self.color.lstrip("#").upper())
        elif self.theme_color is not None:
            solid_fill = etree.SubElement(rpr, qn("a:solidFill"))
            scheme = etree.SubElement(solid_fill, qn("a:schemeClr"))
            scheme.set("val", self.theme_color)

        # Highlight (uses a:highlight element)
        if self.highlight is not None:
            highlight = etree.SubElement(rpr, qn("a:highlight"))
            srgb = etree.SubElement(highlight, qn("a:srgbClr"))
            srgb.set("val", self.highlight.lstrip("#").upper())

        # Font
        if self.font_family is not None:
            latin = etree.SubElement(rpr, qn("a:latin"))
            latin.set("typeface", self.font_family)

        # Hyperlink
        if self.hyperlink is not None:
            # Hyperlink is handled at the run level, not in rPr
            # We store it here but apply it in Run.to_element()
            pass

        return rpr

    @classmethod
    def from_element(cls, elem: etree._Element | None) -> RunProperties:
        """Parse rPr element."""
        if elem is None:
            return cls()

        props = cls()

        # Language
        lang = elem.get("lang")
        if lang:
            props.language = lang

        # Font size (in hundredths of a point)
        sz = elem.get("sz")
        if sz:
            props.font_size = int(sz)

        # Bold, italic, underline
        b = elem.get("b")
        if b is not None:
            props.bold = b == "1"
        i = elem.get("i")
        if i is not None:
            props.italic = i == "1"
        u = elem.get("u")
        if u is not None:
            props.underline = u not in ("none", "0")

        # Strikethrough
        strike = elem.get("strike")
        if strike is not None:
            if strike == "sngStrike":
                props.strikethrough = True
            elif strike == "dblStrike":
                props.double_strikethrough = True
            elif strike == "noStrike":
                props.strikethrough = False
                props.double_strikethrough = False

        # Superscript/Subscript (baseline)
        baseline = elem.get("baseline")
        if baseline is not None:
            baseline_val = int(baseline)
            if baseline_val > 0:
                props.superscript = True
            elif baseline_val < 0:
                props.subscript = True

        # Character spacing
        spc = elem.get("spc")
        if spc:
            props.char_spacing = int(spc)

        # Color
        solid_fill = elem.find(qn("a:solidFill"))
        if solid_fill is not None:
            srgb = solid_fill.find(qn("a:srgbClr"))
            if srgb is not None:
                props.color = srgb.get("val")
            scheme = solid_fill.find(qn("a:schemeClr"))
            if scheme is not None:
                props.theme_color = scheme.get("val")

        # Highlight
        highlight = elem.find(qn("a:highlight"))
        if highlight is not None:
            srgb = highlight.find(qn("a:srgbClr"))
            if srgb is not None:
                props.highlight = srgb.get("val")

        # Font
        latin = elem.find(qn("a:latin"))
        if latin is not None:
            props.font_family = latin.get("typeface")

        return props


@dataclass
class Run:
    """A run of text with consistent formatting."""

    text: str
    properties: RunProperties = field(default_factory=RunProperties)

    def to_element(self) -> etree._Element:
        """Create a:r element."""
        r = etree.Element(qn("a:r"))

        # Run properties
        rpr = self.properties.to_element()
        if rpr is None:
            rpr = etree.Element(qn("a:rPr"))

        # Add hyperlink to run properties if present
        if self.properties.hyperlink is not None:
            hlink = etree.SubElement(rpr, qn("a:hlinkClick"))
            # Set relationship ID for the hyperlink
            hlink.set(qn("r:id"), self.properties.hyperlink)

        # Only append rpr if it has content
        if len(rpr) > 0 or rpr.attrib:
            r.append(rpr)

        # Text content
        t = etree.SubElement(r, qn("a:t"))
        t.text = self.text

        return r

    @classmethod
    def from_element(cls, elem: etree._Element) -> Run:
        """Parse a:r element."""
        text = ""
        t_elem = elem.find(qn("a:t"))
        if t_elem is not None and t_elem.text:
            text = t_elem.text

        rpr = elem.find(qn("a:rPr"))
        props = RunProperties.from_element(rpr)

        # Check for hyperlink
        if rpr is not None:
            hlink = rpr.find(qn("a:hlinkClick"))
            if hlink is not None:
                r_id = hlink.get(qn("r:id"))
                if r_id:
                    props.hyperlink = r_id

        return cls(text=text, properties=props)


@dataclass
class ParagraphProperties:
    """Paragraph-level formatting."""

    level: int = 0  # Bullet level (0 = top level)
    alignment: str | None = None  # "l", "ctr", "r", "just"
    bullet: bool | None = None  # True = bulleted, False = no bullet, None = inherit

    def to_element(self) -> etree._Element:
        """Create pPr element."""
        ppr = etree.Element(qn("a:pPr"))

        if self.level > 0:
            ppr.set("lvl", str(self.level))

        if self.alignment:
            ppr.set("algn", self.alignment)

        if self.bullet is False:
            etree.SubElement(ppr, qn("a:buNone"))

        return ppr

    @classmethod
    def from_element(cls, elem: etree._Element | None) -> ParagraphProperties:
        """Parse pPr element."""
        if elem is None:
            return cls()

        props = cls()

        lvl = elem.get("lvl")
        if lvl:
            props.level = int(lvl)

        props.alignment = elem.get("algn")

        # Check for bullet
        if elem.find(qn("a:buNone")) is not None:
            props.bullet = False
        elif (
            elem.find(qn("a:buChar")) is not None
            or elem.find(qn("a:buAutoNum")) is not None
        ):
            props.bullet = True

        return props


@dataclass
class Paragraph:
    """A paragraph containing runs of text."""

    runs: list[Run] = field(default_factory=list)
    properties: ParagraphProperties = field(default_factory=ParagraphProperties)
    end_para_rpr: RunProperties | None = None  # End paragraph run properties

    @property
    def text(self) -> str:
        """Get full text of paragraph."""
        return "".join(r.text for r in self.runs)

    @text.setter
    def text(self, value: str) -> None:
        """Set text, replacing all runs with a single run."""
        props = self.runs[0].properties if self.runs else RunProperties()
        self.runs = [Run(text=value, properties=props)]

    def to_element(self) -> etree._Element:
        """Create a:p element."""
        p = etree.Element(qn("a:p"))

        # Paragraph properties
        ppr = self.properties.to_element()
        if len(ppr) > 0 or ppr.attrib:
            p.append(ppr)

        # Runs
        for run in self.runs:
            p.append(run.to_element())

        # End paragraph run properties
        if self.end_para_rpr is not None:
            end_rpr = self.end_para_rpr.to_element()
            if end_rpr is not None:
                end_rpr.tag = qn("a:endParaRPr")
                p.append(end_rpr)
        else:
            # Add minimal endParaRPr
            etree.SubElement(p, qn("a:endParaRPr"))

        return p

    @classmethod
    def from_element(cls, elem: etree._Element) -> Paragraph:
        """Parse a:p element."""
        para = cls()

        # Parse properties
        ppr = elem.find(qn("a:pPr"))
        para.properties = ParagraphProperties.from_element(ppr)

        # Parse runs
        for r_elem in elem.findall(qn("a:r")):
            para.runs.append(Run.from_element(r_elem))

        # Also get text from field elements (a:fld)
        for fld in elem.findall(qn("a:fld")):
            t_elem = fld.find(qn("a:t"))
            if t_elem is not None and t_elem.text:
                para.runs.append(Run(text=t_elem.text))

        # End para properties
        end_rpr = elem.find(qn("a:endParaRPr"))
        if end_rpr is not None:
            para.end_para_rpr = RunProperties.from_element(end_rpr)

        return para


@dataclass
class TextBody:
    """A text body (txBody) containing paragraphs."""

    paragraphs: list[Paragraph] = field(default_factory=list)
    body_properties: dict = field(default_factory=dict)

    @property
    def text(self) -> str:
        """Get full text with newlines between paragraphs."""
        return "\n".join(p.text for p in self.paragraphs)

    @text.setter
    def text(self, value: str) -> None:
        """Set text, splitting on newlines into paragraphs."""
        lines = value.split("\n")
        self.paragraphs = [Paragraph(runs=[Run(text=line)]) for line in lines]

    def to_element(self) -> etree._Element:
        """Create txBody element."""
        tx_body = etree.Element(qn("p:txBody"))

        # Body properties
        body_pr = etree.SubElement(tx_body, qn("a:bodyPr"))
        for key, val in self.body_properties.items():
            body_pr.set(key, str(val))

        # List style (empty for now)
        etree.SubElement(tx_body, qn("a:lstStyle"))

        # Paragraphs
        if self.paragraphs:
            for para in self.paragraphs:
                tx_body.append(para.to_element())
        else:
            # Must have at least one paragraph
            tx_body.append(Paragraph().to_element())

        return tx_body

    @classmethod
    def from_element(cls, elem: etree._Element) -> TextBody:
        """Parse txBody element."""
        tb = cls()

        # Body properties
        body_pr = elem.find(qn("a:bodyPr"))
        if body_pr is not None:
            tb.body_properties = dict(body_pr.attrib)

        # Paragraphs
        for p_elem in elem.findall(qn("a:p")):
            tb.paragraphs.append(Paragraph.from_element(p_elem))

        return tb


def create_text_body(
    text: str,
    font_size: int | None = None,
    font_family: str | None = None,
    bold: bool = False,
    color: str | None = None,
) -> etree._Element:
    """Create a simple text body element.

    Args:
        text: The text content
        font_size: Font size in points (will be converted to centipoints)
        font_family: Font family name
        bold: Whether text should be bold
        color: Hex color string (with or without #)

    Returns:
        txBody element
    """
    props = RunProperties(
        font_size=font_size * 100 if font_size else None,
        font_family=font_family,
        bold=bold if bold else None,
        color=color.lstrip("#") if color else None,
    )

    run = Run(text=text, properties=props)
    para = Paragraph(runs=[run])
    body = TextBody(paragraphs=[para])

    return body.to_element()


def create_bullet_body(
    items: list[str],
    levels: list[int] | None = None,
    font_size: int | None = None,
    font_family: str | None = None,
) -> etree._Element:
    """Create a text body with bullet points.

    Args:
        items: List of bullet point texts
        levels: Optional list of indent levels (0-8)
        font_size: Font size in points
        font_family: Font family name

    Returns:
        txBody element
    """
    if levels is None:
        levels = [0] * len(items)

    paragraphs = []
    for item, level in zip(items, levels, strict=False):
        props = RunProperties(
            font_size=font_size * 100 if font_size else None,
            font_family=font_family,
        )
        run = Run(text=item, properties=props)
        para_props = ParagraphProperties(level=level)
        para = Paragraph(runs=[run], properties=para_props)
        paragraphs.append(para)

    body = TextBody(paragraphs=paragraphs)
    return body.to_element()
