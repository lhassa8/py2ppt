"""Slide XML part handling.

Each slide is stored as ppt/slides/slideN.xml with relationships
in ppt/slides/_rels/slideN.xml.rels
"""

from __future__ import annotations

from lxml import etree

from .ns import CONTENT_TYPE, REL_TYPE, nsmap, qn
from .package import Package
from .shapes import Picture, Shape, ShapeTree, Table


class SlidePart:
    """Handles a slide XML part (ppt/slides/slideN.xml).

    A slide contains:
    - cSld (common slide data)
      - spTree (shape tree)
    - clrMapOvr (color map override - optional)
    """

    def __init__(self, element: etree._Element) -> None:
        self._element = element
        self._shape_tree: ShapeTree | None = None

    @property
    def element(self) -> etree._Element:
        return self._element

    @property
    def shape_tree(self) -> ShapeTree:
        """Get the shape tree, parsing if necessary."""
        if self._shape_tree is None:
            sp_tree_elem = self._element.find(f".//{qn('p:spTree')}")
            if sp_tree_elem is not None:
                self._shape_tree = ShapeTree.from_element(sp_tree_elem)
            else:
                self._shape_tree = ShapeTree()
        return self._shape_tree

    def get_placeholder(
        self, ph_type: str | None = None, ph_idx: int | None = None
    ) -> Shape | None:
        """Find a placeholder by type and/or index."""
        return self.shape_tree.get_placeholder(ph_type, ph_idx)

    def get_title_placeholder(self) -> Shape | None:
        """Get the title placeholder."""
        # Try various title types
        for ph_type in ["title", "ctrTitle"]:
            ph = self.get_placeholder(ph_type=ph_type)
            if ph:
                return ph
        return None

    def get_body_placeholder(self) -> Shape | None:
        """Get the body/content placeholder."""
        for ph_type in ["body", "obj"]:
            ph = self.get_placeholder(ph_type=ph_type)
            if ph:
                return ph
        return None

    def get_subtitle_placeholder(self) -> Shape | None:
        """Get the subtitle placeholder."""
        return self.get_placeholder(ph_type="subTitle")

    def add_shape(self, shape: Shape) -> None:
        """Add a shape to the slide."""
        self.shape_tree.add_shape(shape)

    def add_picture(self, pic: Picture) -> None:
        """Add a picture to the slide."""
        self.shape_tree.add_shape(pic)

    def add_table(self, table: Table) -> None:
        """Add a table to the slide."""
        self.shape_tree.add_shape(table)

    def update_xml(self) -> None:
        """Update the XML element from the shape tree."""
        c_sld = self._element.find(qn("p:cSld"))
        if c_sld is None:
            c_sld = etree.SubElement(self._element, qn("p:cSld"))

        # Remove existing spTree
        existing = c_sld.find(qn("p:spTree"))
        if existing is not None:
            c_sld.remove(existing)

        # Add updated spTree
        c_sld.append(self.shape_tree.to_element())

    def to_xml(self) -> bytes:
        """Serialize to XML bytes."""
        self.update_xml()
        return etree.tostring(
            self._element,
            xml_declaration=True,
            encoding="UTF-8",
            standalone=True,
        )

    @classmethod
    def from_xml(cls, xml_bytes: bytes) -> SlidePart:
        """Parse from XML bytes."""
        element = etree.fromstring(xml_bytes)
        return cls(element)

    @classmethod
    def new(cls) -> SlidePart:
        """Create a new blank slide."""
        nsmap_slide = {
            None: nsmap["p"],
            "a": nsmap["a"],
            "r": nsmap["r"],
            "p14": nsmap["p14"],
        }

        root = etree.Element(qn("p:sld"), nsmap=nsmap_slide)

        # Common slide data
        c_sld = etree.SubElement(root, qn("p:cSld"))

        # Empty shape tree
        sp_tree = ShapeTree()
        c_sld.append(sp_tree.to_element())

        # Color map override
        clr_map_ovr = etree.SubElement(root, qn("p:clrMapOvr"))
        etree.SubElement(clr_map_ovr, qn("a:masterClrMapping"))

        return cls(root)


def get_slide_part(pkg: Package, slide_number: int) -> SlidePart | None:
    """Get a slide part by slide number (1-indexed).

    This uses the presentation.xml to find the actual slide path.
    """
    from .presentation import get_presentation_part

    pres_part = get_presentation_part(pkg)
    if pres_part is None:
        return None

    slide_refs = pres_part.get_slide_refs()
    if slide_number < 1 or slide_number > len(slide_refs):
        return None

    ref = slide_refs[slide_number - 1]

    # Get slide path from relationship
    pres_rels = pkg.get_part_rels("ppt/presentation.xml")
    rel = pres_rels.get(ref.r_id)
    if rel is None:
        return None

    # Relationship target is relative to ppt/
    if rel.target.startswith("/"):
        slide_path = rel.target.lstrip("/")
    else:
        slide_path = f"ppt/{rel.target}"

    content = pkg.get_part(slide_path)
    if content is None:
        return None

    return SlidePart.from_xml(content)


def add_slide_to_package(
    pkg: Package,
    slide_part: SlidePart,
    layout_r_id: str,
    position: int | None = None,
) -> int:
    """Add a slide to the package.

    Args:
        pkg: The package to modify
        slide_part: The slide to add
        layout_r_id: Relationship ID of the layout to use (from master)
        position: Insert position (0-indexed). None = append at end.

    Returns:
        The slide number of the new slide (1-indexed).
    """
    from .presentation import get_presentation_part

    pres_part = get_presentation_part(pkg)
    if pres_part is None:
        raise ValueError("Package has no presentation part")

    # Determine slide number
    existing_refs = pres_part.get_slide_refs()
    slide_num = len(existing_refs) + 1

    slide_path = f"ppt/slides/slide{slide_num}.xml"

    # Add slide part
    pkg.set_part(slide_path, slide_part.to_xml(), CONTENT_TYPE.SLIDE)

    # Create slide relationships
    slide_rels = pkg.get_part_rels(slide_path)
    slide_rels.add(
        rel_type=REL_TYPE.SLIDE_LAYOUT,
        target=f"../slideLayouts/slideLayout{layout_r_id.replace('rId', '')}.xml",
        r_id="rId1",
    )
    pkg.set_part_rels(slide_path, slide_rels)

    # Add slide reference to presentation
    pres_rels = pkg.get_part_rels("ppt/presentation.xml")
    r_id = pres_rels.add(
        rel_type=REL_TYPE.SLIDE,
        target=f"slides/slide{slide_num}.xml",
    )

    pres_part.add_slide_ref(r_id, position)

    # Update presentation part
    pkg.set_part(
        "ppt/presentation.xml",
        pres_part.to_xml(),
        CONTENT_TYPE.PRESENTATION,
    )

    return slide_num


def update_slide_in_package(
    pkg: Package, slide_number: int, slide_part: SlidePart
) -> bool:
    """Update a slide in the package.

    Args:
        pkg: The package to modify
        slide_number: The slide number (1-indexed)
        slide_part: The updated slide

    Returns:
        True if successful, False if slide not found.
    """
    from .presentation import get_presentation_part

    pres_part = get_presentation_part(pkg)
    if pres_part is None:
        return False

    slide_refs = pres_part.get_slide_refs()
    if slide_number < 1 or slide_number > len(slide_refs):
        return False

    ref = slide_refs[slide_number - 1]

    # Get slide path from relationship
    pres_rels = pkg.get_part_rels("ppt/presentation.xml")
    rel = pres_rels.get(ref.r_id)
    if rel is None:
        return False

    if rel.target.startswith("/"):
        slide_path = rel.target.lstrip("/")
    else:
        slide_path = f"ppt/{rel.target}"

    pkg.set_part(slide_path, slide_part.to_xml(), CONTENT_TYPE.SLIDE)
    return True


def remove_slide_from_package(pkg: Package, slide_number: int) -> bool:
    """Remove a slide from the package.

    Args:
        pkg: The package to modify
        slide_number: The slide number (1-indexed)

    Returns:
        True if successful, False if slide not found.
    """
    from .presentation import get_presentation_part

    pres_part = get_presentation_part(pkg)
    if pres_part is None:
        return False

    slide_refs = pres_part.get_slide_refs()
    if slide_number < 1 or slide_number > len(slide_refs):
        return False

    ref = slide_refs[slide_number - 1]

    # Get slide path
    pres_rels = pkg.get_part_rels("ppt/presentation.xml")
    rel = pres_rels.get(ref.r_id)
    if rel is None:
        return False

    if rel.target.startswith("/"):
        slide_path = rel.target.lstrip("/")
    else:
        slide_path = f"ppt/{rel.target}"

    # Remove slide reference from presentation
    pres_part.remove_slide_ref(ref.r_id)
    pres_rels.remove(ref.r_id)

    # Update presentation
    pkg.set_part(
        "ppt/presentation.xml",
        pres_part.to_xml(),
        CONTENT_TYPE.PRESENTATION,
    )

    # Remove slide part
    pkg.remove_part(slide_path)

    return True
