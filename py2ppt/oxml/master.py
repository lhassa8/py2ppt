"""Slide master XML part handling.

Slide masters define the overall look and feel of slides,
stored in ppt/slideMasters/slideMasterN.xml.
"""

from __future__ import annotations

from dataclasses import dataclass, field
from typing import Dict, List, Optional, Tuple

from lxml import etree

from .ns import CONTENT_TYPE, REL_TYPE, nsmap, qn
from .package import Package
from .shapes import ShapeTree


@dataclass
class MasterLayoutRef:
    """Reference to a layout from the master."""

    r_id: str
    layout_index: int
    layout_name: str


class SlideMasterPart:
    """Handles a slide master XML part.

    A master contains:
    - cSld (common slide data)
      - bg (background)
      - spTree (shape tree with default placeholders)
    - clrMap (color map)
    - sldLayoutIdLst (list of associated layouts)
    - txStyles (text styles)
    """

    def __init__(self, element: etree._Element) -> None:
        self._element = element
        self._shape_tree: Optional[ShapeTree] = None

    @property
    def element(self) -> etree._Element:
        return self._element

    @property
    def shape_tree(self) -> ShapeTree:
        """Get the shape tree."""
        if self._shape_tree is None:
            sp_tree_elem = self._element.find(f".//{qn('p:spTree')}")
            if sp_tree_elem is not None:
                self._shape_tree = ShapeTree.from_element(sp_tree_elem)
            else:
                self._shape_tree = ShapeTree()
        return self._shape_tree

    def get_layout_refs(self) -> List[Tuple[str, int]]:
        """Get layout references as (rId, layoutId) tuples."""
        refs = []
        sld_layout_id_lst = self._element.find(qn("p:sldLayoutIdLst"))
        if sld_layout_id_lst is not None:
            for layout_id in sld_layout_id_lst.findall(qn("p:sldLayoutId")):
                r_id = layout_id.get(qn("r:id"), "")
                l_id = int(layout_id.get("id", "0"))
                refs.append((r_id, l_id))
        return refs

    def get_color_map(self) -> Dict[str, str]:
        """Get the color map (scheme -> actual color role mapping)."""
        color_map = {}
        clr_map = self._element.find(qn("p:clrMap"))
        if clr_map is not None:
            for key in [
                "bg1", "tx1", "bg2", "tx2",
                "accent1", "accent2", "accent3", "accent4",
                "accent5", "accent6", "hlink", "folHlink",
            ]:
                val = clr_map.get(key)
                if val:
                    color_map[key] = val
        return color_map

    def to_xml(self) -> bytes:
        """Serialize to XML bytes."""
        return etree.tostring(
            self._element,
            xml_declaration=True,
            encoding="UTF-8",
            standalone=True,
        )

    @classmethod
    def from_xml(cls, xml_bytes: bytes) -> "SlideMasterPart":
        """Parse from XML bytes."""
        element = etree.fromstring(xml_bytes)
        return cls(element)


def get_master_parts(pkg: Package) -> Dict[str, SlideMasterPart]:
    """Get all slide master parts from the package.

    Returns:
        Dict mapping master path to SlideMasterPart
    """
    masters = {}

    for part_name, content in pkg.iter_parts():
        if part_name.startswith("ppt/slideMasters/") and part_name.endswith(".xml"):
            if "/_rels/" not in part_name:
                masters[part_name] = SlideMasterPart.from_xml(content)

    return masters


def get_primary_master(pkg: Package) -> Optional[SlideMasterPart]:
    """Get the primary (first) slide master.

    Most presentations have a single master, so this is often sufficient.
    """
    masters = get_master_parts(pkg)
    if not masters:
        return None

    # Return the first one (by filename order)
    sorted_paths = sorted(masters.keys())
    return masters[sorted_paths[0]]


def get_master_for_layout(
    pkg: Package, layout_index: int
) -> Optional[SlideMasterPart]:
    """Find the master that owns a specific layout.

    Args:
        pkg: The package
        layout_index: The layout index (0-indexed)

    Returns:
        The SlideMasterPart that contains the layout, or None
    """
    masters = get_master_parts(pkg)

    for master_path, master in masters.items():
        master_rels = pkg.get_part_rels(master_path)

        for r_id, _ in master.get_layout_refs():
            rel = master_rels.get(r_id)
            if rel:
                # Extract layout number from target path
                # Target is like "../slideLayouts/slideLayout1.xml"
                target = rel.target
                if "slideLayout" in target:
                    try:
                        num = int(
                            target.split("slideLayout")[1].split(".")[0]
                        )
                        if num - 1 == layout_index:  # Layout numbers are 1-based
                            return master
                    except (ValueError, IndexError):
                        pass

    return None


def create_minimal_master() -> SlideMasterPart:
    """Create a minimal slide master part.

    This creates a basic master with:
    - Default placeholders for title and body
    - Basic color map
    - Empty layout list
    """
    nsmap_master = {
        None: nsmap["p"],
        "a": nsmap["a"],
        "r": nsmap["r"],
    }

    root = etree.Element(qn("p:sldMaster"), nsmap=nsmap_master)

    # Common slide data
    c_sld = etree.SubElement(root, qn("p:cSld"))

    # Background
    bg = etree.SubElement(c_sld, qn("p:bg"))
    bg_ref = etree.SubElement(bg, qn("p:bgRef"))
    bg_ref.set("idx", "1001")
    scheme_clr = etree.SubElement(bg_ref, qn("a:schemeClr"))
    scheme_clr.set("val", "bg1")

    # Empty shape tree
    sp_tree = ShapeTree()
    c_sld.append(sp_tree.to_element())

    # Color map
    clr_map = etree.SubElement(root, qn("p:clrMap"))
    clr_map.set("bg1", "lt1")
    clr_map.set("tx1", "dk1")
    clr_map.set("bg2", "lt2")
    clr_map.set("tx2", "dk2")
    clr_map.set("accent1", "accent1")
    clr_map.set("accent2", "accent2")
    clr_map.set("accent3", "accent3")
    clr_map.set("accent4", "accent4")
    clr_map.set("accent5", "accent5")
    clr_map.set("accent6", "accent6")
    clr_map.set("hlink", "hlink")
    clr_map.set("folHlink", "folHlink")

    # Empty layout ID list
    etree.SubElement(root, qn("p:sldLayoutIdLst"))

    return SlideMasterPart(root)
