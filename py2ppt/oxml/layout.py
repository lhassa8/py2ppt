"""Slide layout XML part handling.

Slide layouts define the structure and default content of slides,
stored in ppt/slideLayouts/slideLayoutN.xml.
"""

from __future__ import annotations

from dataclasses import dataclass, field
from typing import Dict, List, Optional, Tuple

from lxml import etree

from .ns import CONTENT_TYPE, REL_TYPE, nsmap, qn
from .package import Package
from .shapes import PlaceholderInfo, Position, Shape, ShapeTree


@dataclass
class LayoutPlaceholder:
    """Information about a placeholder in a layout."""

    type: str  # Placeholder type (title, body, etc.)
    idx: Optional[int]  # Placeholder index
    name: str  # Shape name
    position: Position  # Position and size


@dataclass
class LayoutInfo:
    """Information about a slide layout."""

    name: str
    index: int  # Index in the layouts list
    placeholders: List[LayoutPlaceholder] = field(default_factory=list)
    master_r_id: str = ""  # Relationship to slide master


class SlideLayoutPart:
    """Handles a slide layout XML part.

    A layout contains:
    - cSld (common slide data)
      - spTree (shape tree with placeholders)
    - clrMapOvr (color map override - optional)
    """

    def __init__(self, element: etree._Element) -> None:
        self._element = element
        self._shape_tree: Optional[ShapeTree] = None

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

    def get_name(self) -> str:
        """Get the layout name from cSld/@name."""
        c_sld = self._element.find(qn("p:cSld"))
        if c_sld is not None:
            return c_sld.get("name", "")
        return ""

    def get_type(self) -> Optional[str]:
        """Get the layout type if present."""
        return self._element.get("type")

    def get_placeholders(self) -> List[LayoutPlaceholder]:
        """Get all placeholder information from this layout."""
        placeholders = []

        for shape in self.shape_tree.get_placeholders():
            if shape.placeholder:
                placeholders.append(
                    LayoutPlaceholder(
                        type=shape.placeholder.type or "body",
                        idx=shape.placeholder.idx,
                        name=shape.name,
                        position=shape.position,
                    )
                )

        return placeholders

    def get_placeholder_by_type(
        self, ph_type: str
    ) -> Optional[LayoutPlaceholder]:
        """Get a placeholder by type."""
        for ph in self.get_placeholders():
            if ph.type == ph_type:
                return ph
        return None

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
    def from_xml(cls, xml_bytes: bytes) -> "SlideLayoutPart":
        """Parse from XML bytes."""
        element = etree.fromstring(xml_bytes)
        return cls(element)

    @classmethod
    def new(cls, name: str = "Custom Layout") -> "SlideLayoutPart":
        """Create a new blank layout."""
        nsmap_layout = {
            None: nsmap["p"],
            "a": nsmap["a"],
            "r": nsmap["r"],
        }

        root = etree.Element(qn("p:sldLayout"), nsmap=nsmap_layout)

        # Common slide data
        c_sld = etree.SubElement(root, qn("p:cSld"))
        c_sld.set("name", name)

        # Empty shape tree
        sp_tree = ShapeTree()
        c_sld.append(sp_tree.to_element())

        # Color map override
        clr_map_ovr = etree.SubElement(root, qn("p:clrMapOvr"))
        etree.SubElement(clr_map_ovr, qn("a:masterClrMapping"))

        return cls(root)


def get_layout_parts(pkg: Package) -> Dict[str, SlideLayoutPart]:
    """Get all slide layout parts from the package.

    Returns:
        Dict mapping layout path to SlideLayoutPart
    """
    layouts = {}

    for part_name, content in pkg.iter_parts():
        if part_name.startswith("ppt/slideLayouts/") and part_name.endswith(".xml"):
            if "/_rels/" not in part_name:
                layouts[part_name] = SlideLayoutPart.from_xml(content)

    return layouts


def _extract_layout_num(path: str) -> int:
    """Extract layout number from a path like 'ppt/slideLayouts/slideLayout1.xml'."""
    filename = path.split("/")[-1]  # Get just the filename
    num_str = filename.replace("slideLayout", "").replace(".xml", "")
    return int(num_str) if num_str.isdigit() else 0


def get_layout_info_list(pkg: Package) -> List[LayoutInfo]:
    """Get information about all layouts in the package.

    Returns:
        List of LayoutInfo objects, ordered by layout index.
    """
    layouts = []

    # Get all layout parts
    layout_parts = get_layout_parts(pkg)
    sorted_paths = sorted(layout_parts.keys(), key=_extract_layout_num)

    for idx, path in enumerate(sorted_paths):
        part = layout_parts[path]
        placeholders = part.get_placeholders()

        # Get master relationship
        layout_rels = pkg.get_part_rels(path)
        master_rels = layout_rels.find_by_type(REL_TYPE.SLIDE_MASTER)
        master_r_id = master_rels[0].r_id if master_rels else ""

        layouts.append(
            LayoutInfo(
                name=part.get_name() or f"Layout {idx + 1}",
                index=idx,
                placeholders=placeholders,
                master_r_id=master_r_id,
            )
        )

    return layouts


def get_layout_by_name(
    pkg: Package, name: str, fuzzy: bool = True
) -> Optional[Tuple[SlideLayoutPart, int]]:
    """Find a layout by name.

    Args:
        pkg: The package to search
        name: Layout name to find
        fuzzy: If True, do fuzzy matching (case-insensitive, partial)

    Returns:
        Tuple of (SlideLayoutPart, index) or None if not found
    """
    layout_parts = get_layout_parts(pkg)
    sorted_paths = sorted(layout_parts.keys(), key=_extract_layout_num)

    name_lower = name.lower().replace("_", " ").replace("-", " ")

    for idx, path in enumerate(sorted_paths):
        part = layout_parts[path]
        layout_name = part.get_name()

        if fuzzy:
            layout_name_lower = layout_name.lower().replace("_", " ").replace("-", " ")
            # Exact match (case insensitive)
            if layout_name_lower == name_lower:
                return (part, idx)
            # Partial match
            if name_lower in layout_name_lower or layout_name_lower in name_lower:
                return (part, idx)
        else:
            if layout_name == name:
                return (part, idx)

    return None


def get_layout_by_index(pkg: Package, index: int) -> Optional[SlideLayoutPart]:
    """Get a layout by index (0-indexed).

    Args:
        pkg: The package to search
        index: Layout index

    Returns:
        SlideLayoutPart or None if not found
    """
    layout_parts = get_layout_parts(pkg)
    sorted_paths = sorted(layout_parts.keys(), key=_extract_layout_num)

    if 0 <= index < len(sorted_paths):
        return layout_parts[sorted_paths[index]]

    return None


# Common layout names and their likely types
LAYOUT_NAME_PATTERNS = {
    "title slide": ["title", "ctrTitle", "subTitle"],
    "title only": ["title"],
    "title and content": ["title", "body"],
    "section header": ["title", "body"],
    "two content": ["title", "body", "body"],
    "comparison": ["title", "body", "body"],
    "content with caption": ["title", "body", "body"],
    "picture with caption": ["title", "body", "pic"],
    "blank": [],
}


def normalize_layout_name(name: str) -> str:
    """Normalize a layout name for matching.

    Examples:
        "Title Slide" -> "title slide"
        "title_slide" -> "title slide"
        "Title and Content" -> "title and content"
    """
    return name.lower().replace("_", " ").replace("-", " ").strip()


def guess_layout_type(layout_name: str) -> List[str]:
    """Guess placeholder types for a layout name.

    Returns list of expected placeholder types.
    """
    normalized = normalize_layout_name(layout_name)

    for pattern, types in LAYOUT_NAME_PATTERNS.items():
        if pattern in normalized or normalized in pattern:
            return types

    # Default: assume title and body
    return ["title", "body"]
