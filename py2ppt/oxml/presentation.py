"""Presentation XML part handling.

The presentation.xml file is the main entry point for a PowerPoint file,
containing references to slides, masters, and presentation settings.
"""

from __future__ import annotations

from dataclasses import dataclass
from typing import Dict, List, Optional, Tuple

from lxml import etree

from .ns import CONTENT_TYPE, REL_TYPE, nsmap, qn
from .package import Package, RelationshipCollection


@dataclass
class SlideRef:
    """Reference to a slide in the presentation."""

    r_id: str
    slide_id: int


class PresentationPart:
    """Handles the presentation.xml part.

    The presentation part contains:
    - sldIdLst: List of slides with IDs and relationship references
    - sldMasterIdLst: List of slide masters
    - sldSz: Slide size
    - notesSz: Notes size
    """

    PART_NAME = "ppt/presentation.xml"

    def __init__(self, element: etree._Element) -> None:
        self._element = element
        self._next_slide_id = 256  # PowerPoint starts slide IDs at 256

    @property
    def element(self) -> etree._Element:
        return self._element

    def get_slide_refs(self) -> List[SlideRef]:
        """Get all slide references in order."""
        refs = []
        sld_id_lst = self._element.find(qn("p:sldIdLst"))
        if sld_id_lst is not None:
            for sld_id in sld_id_lst.findall(qn("p:sldId")):
                r_id = sld_id.get(qn("r:id"), "")
                slide_id = int(sld_id.get("id", "0"))
                refs.append(SlideRef(r_id=r_id, slide_id=slide_id))
                if slide_id >= self._next_slide_id:
                    self._next_slide_id = slide_id + 1
        return refs

    def get_master_refs(self) -> List[Tuple[str, int]]:
        """Get slide master references as (rId, masterId) tuples."""
        refs = []
        master_id_lst = self._element.find(qn("p:sldMasterIdLst"))
        if master_id_lst is not None:
            for master_id in master_id_lst.findall(qn("p:sldMasterId")):
                r_id = master_id.get(qn("r:id"), "")
                m_id = int(master_id.get("id", "0"))
                refs.append((r_id, m_id))
        return refs

    def add_slide_ref(self, r_id: str, position: Optional[int] = None) -> int:
        """Add a slide reference and return its slide ID."""
        sld_id_lst = self._element.find(qn("p:sldIdLst"))
        if sld_id_lst is None:
            # Create sldIdLst after sldMasterIdLst
            master_lst = self._element.find(qn("p:sldMasterIdLst"))
            sld_id_lst = etree.Element(qn("p:sldIdLst"))
            if master_lst is not None:
                idx = list(self._element).index(master_lst) + 1
                self._element.insert(idx, sld_id_lst)
            else:
                self._element.insert(0, sld_id_lst)

        slide_id = self._next_slide_id
        self._next_slide_id += 1

        sld_id_elem = etree.Element(
            qn("p:sldId"),
            attrib={
                "id": str(slide_id),
                qn("r:id"): r_id,
            },
        )

        if position is not None and 0 <= position < len(sld_id_lst):
            sld_id_lst.insert(position, sld_id_elem)
        else:
            sld_id_lst.append(sld_id_elem)

        return slide_id

    def remove_slide_ref(self, r_id: str) -> bool:
        """Remove a slide reference by relationship ID."""
        sld_id_lst = self._element.find(qn("p:sldIdLst"))
        if sld_id_lst is None:
            return False

        for sld_id in sld_id_lst.findall(qn("p:sldId")):
            if sld_id.get(qn("r:id")) == r_id:
                sld_id_lst.remove(sld_id)
                return True
        return False

    def reorder_slides(self, r_id_order: List[str]) -> None:
        """Reorder slides according to the given relationship ID order."""
        sld_id_lst = self._element.find(qn("p:sldIdLst"))
        if sld_id_lst is None:
            return

        # Build map of r_id -> element
        elements = {}
        for sld_id in list(sld_id_lst):
            r_id = sld_id.get(qn("r:id"), "")
            elements[r_id] = sld_id
            sld_id_lst.remove(sld_id)

        # Re-add in new order
        for r_id in r_id_order:
            if r_id in elements:
                sld_id_lst.append(elements[r_id])

    def get_slide_size(self) -> Tuple[int, int]:
        """Get slide size in EMUs as (width, height)."""
        sld_sz = self._element.find(qn("p:sldSz"))
        if sld_sz is not None:
            cx = int(sld_sz.get("cx", "9144000"))
            cy = int(sld_sz.get("cy", "6858000"))
            return (cx, cy)
        return (9144000, 6858000)  # Default 10" x 7.5"

    def set_slide_size(self, width: int, height: int) -> None:
        """Set slide size in EMUs."""
        sld_sz = self._element.find(qn("p:sldSz"))
        if sld_sz is None:
            sld_sz = etree.SubElement(self._element, qn("p:sldSz"))
        sld_sz.set("cx", str(width))
        sld_sz.set("cy", str(height))
        sld_sz.set("type", "custom")

    def to_xml(self) -> bytes:
        """Serialize to XML bytes."""
        return etree.tostring(
            self._element,
            xml_declaration=True,
            encoding="UTF-8",
            standalone=True,
        )

    @classmethod
    def from_xml(cls, xml_bytes: bytes) -> "PresentationPart":
        """Parse from XML bytes."""
        element = etree.fromstring(xml_bytes)
        return cls(element)

    @classmethod
    def new(cls) -> "PresentationPart":
        """Create a new blank presentation part."""
        nsmap_pres = {
            None: nsmap["p"],
            "a": nsmap["a"],
            "r": nsmap["r"],
            "p14": nsmap["p14"],
        }

        root = etree.Element(qn("p:presentation"), nsmap=nsmap_pres)

        # Add slide masters list (empty for now)
        etree.SubElement(root, qn("p:sldMasterIdLst"))

        # Add empty slide list
        etree.SubElement(root, qn("p:sldIdLst"))

        # Add slide size (standard 10" x 7.5" at 914400 EMU/inch)
        sld_sz = etree.SubElement(root, qn("p:sldSz"))
        sld_sz.set("cx", "9144000")
        sld_sz.set("cy", "6858000")
        sld_sz.set("type", "screen4x3")

        # Notes size
        notes_sz = etree.SubElement(root, qn("p:notesSz"))
        notes_sz.set("cx", "6858000")
        notes_sz.set("cy", "9144000")

        return cls(root)


def setup_presentation_part(pkg: Package, pres_part: PresentationPart) -> None:
    """Set up presentation part in package with proper content type."""
    pkg.set_part(
        PresentationPart.PART_NAME,
        pres_part.to_xml(),
        CONTENT_TYPE.PRESENTATION,
    )


def get_presentation_part(pkg: Package) -> Optional[PresentationPart]:
    """Get presentation part from package."""
    content = pkg.get_part(PresentationPart.PART_NAME)
    if content is None:
        return None
    return PresentationPart.from_xml(content)
