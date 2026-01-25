"""Notes slide XML handling.

Speaker notes are stored in ppt/notesSlides/notesSlideN.xml
with relationships linking them to their parent slides.
"""

from __future__ import annotations

from lxml import etree

from .ns import CONTENT_TYPE, REL_TYPE, nsmap, qn
from .package import Package  # noqa: F401 (Package used for type hints)


class NotesSlidePart:
    """Handles a notes slide XML part."""

    def __init__(self, element: etree._Element) -> None:
        self._element = element

    @property
    def element(self) -> etree._Element:
        return self._element

    def get_text(self) -> str:
        """Get all text from the notes slide."""
        texts = []

        # Find the notes text frame (body placeholder)
        for sp in self._element.findall(f".//{qn('p:sp')}"):
            # Check if this is the body/notes placeholder
            ph = sp.find(f".//{qn('p:ph')}")
            if ph is not None:
                ph_type = ph.get("type", "")
                if ph_type in ("body", ""):
                    # Get text from this shape
                    for t in sp.findall(f".//{qn('a:t')}"):
                        if t.text:
                            texts.append(t.text)

        return "\n".join(texts)

    def set_text(self, text: str) -> None:
        """Set the notes text, replacing any existing content."""
        # Find the notes text frame
        for sp in self._element.findall(f".//{qn('p:sp')}"):
            ph = sp.find(f".//{qn('p:ph')}")
            if ph is not None:
                ph_type = ph.get("type", "")
                if ph_type in ("body", ""):
                    # Find or create txBody
                    tx_body = sp.find(qn("p:txBody"))
                    if tx_body is None:
                        tx_body = etree.SubElement(sp, qn("p:txBody"))
                        etree.SubElement(tx_body, qn("a:bodyPr"))
                        etree.SubElement(tx_body, qn("a:lstStyle"))

                    # Remove existing paragraphs
                    for p in tx_body.findall(qn("a:p")):
                        tx_body.remove(p)

                    # Add new paragraphs
                    for line in text.split("\n"):
                        p = etree.SubElement(tx_body, qn("a:p"))
                        r = etree.SubElement(p, qn("a:r"))
                        t = etree.SubElement(r, qn("a:t"))
                        t.text = line

                    return

    def append_text(self, text: str) -> None:
        """Append text to existing notes."""
        current = self.get_text()
        if current:
            self.set_text(current + "\n" + text)
        else:
            self.set_text(text)

    def to_xml(self) -> bytes:
        """Serialize to XML bytes."""
        return etree.tostring(
            self._element,
            xml_declaration=True,
            encoding="UTF-8",
            standalone=True,
        )

    @classmethod
    def from_xml(cls, xml_bytes: bytes) -> NotesSlidePart:
        """Parse from XML bytes."""
        element = etree.fromstring(xml_bytes)
        return cls(element)

    @classmethod
    def new(cls) -> NotesSlidePart:
        """Create a new blank notes slide."""
        nsmap_notes = {
            None: nsmap["p"],
            "a": nsmap["a"],
            "r": nsmap["r"],
        }

        root = etree.Element(qn("p:notes"), nsmap=nsmap_notes)

        # Common slide data
        c_sld = etree.SubElement(root, qn("p:cSld"))

        # Shape tree
        sp_tree = etree.SubElement(c_sld, qn("p:spTree"))

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

        # Slide image placeholder
        sp1 = etree.SubElement(sp_tree, qn("p:sp"))
        nv_sp_pr1 = etree.SubElement(sp1, qn("p:nvSpPr"))
        c_nv_pr1 = etree.SubElement(nv_sp_pr1, qn("p:cNvPr"))
        c_nv_pr1.set("id", "2")
        c_nv_pr1.set("name", "Slide Image Placeholder")
        etree.SubElement(nv_sp_pr1, qn("p:cNvSpPr"))
        nv_pr1 = etree.SubElement(nv_sp_pr1, qn("p:nvPr"))
        ph1 = etree.SubElement(nv_pr1, qn("p:ph"))
        ph1.set("type", "sldImg")

        sp_pr1 = etree.SubElement(sp1, qn("p:spPr"))
        xfrm1 = etree.SubElement(sp_pr1, qn("a:xfrm"))
        off1 = etree.SubElement(xfrm1, qn("a:off"))
        off1.set("x", "685800")
        off1.set("y", "1143000")
        ext1 = etree.SubElement(xfrm1, qn("a:ext"))
        ext1.set("cx", "5486400")
        ext1.set("cy", "3086100")
        etree.SubElement(sp_pr1, qn("a:noFill"))

        # Notes text placeholder
        sp2 = etree.SubElement(sp_tree, qn("p:sp"))
        nv_sp_pr2 = etree.SubElement(sp2, qn("p:nvSpPr"))
        c_nv_pr2 = etree.SubElement(nv_sp_pr2, qn("p:cNvPr"))
        c_nv_pr2.set("id", "3")
        c_nv_pr2.set("name", "Notes Placeholder")
        etree.SubElement(nv_sp_pr2, qn("p:cNvSpPr"))
        nv_pr2 = etree.SubElement(nv_sp_pr2, qn("p:nvPr"))
        ph2 = etree.SubElement(nv_pr2, qn("p:ph"))
        ph2.set("type", "body")
        ph2.set("idx", "1")

        sp_pr2 = etree.SubElement(sp2, qn("p:spPr"))
        xfrm2 = etree.SubElement(sp_pr2, qn("a:xfrm"))
        off2 = etree.SubElement(xfrm2, qn("a:off"))
        off2.set("x", "685800")
        off2.set("y", "4400550")
        ext2 = etree.SubElement(xfrm2, qn("a:ext"))
        ext2.set("cx", "5486400")
        ext2.set("cy", "3600450")

        # Text body for notes
        tx_body = etree.SubElement(sp2, qn("p:txBody"))
        etree.SubElement(tx_body, qn("a:bodyPr"))
        etree.SubElement(tx_body, qn("a:lstStyle"))
        p = etree.SubElement(tx_body, qn("a:p"))
        etree.SubElement(p, qn("a:endParaRPr"))

        # Color map override
        clr_map_ovr = etree.SubElement(root, qn("p:clrMapOvr"))
        etree.SubElement(clr_map_ovr, qn("a:masterClrMapping"))

        return cls(root)


def get_notes_slide(pkg: Package, slide_number: int) -> NotesSlidePart | None:
    """Get the notes slide for a given slide number."""
    from .presentation import get_presentation_part

    pres_part = get_presentation_part(pkg)
    if pres_part is None:
        return None

    slide_refs = pres_part.get_slide_refs()
    if slide_number < 1 or slide_number > len(slide_refs):
        return None

    ref = slide_refs[slide_number - 1]

    # Get slide path
    pres_rels = pkg.get_part_rels("ppt/presentation.xml")
    rel = pres_rels.get(ref.r_id)
    if rel is None:
        return None

    if rel.target.startswith("/"):
        slide_path = rel.target.lstrip("/")
    else:
        slide_path = f"ppt/{rel.target}"

    # Get notes relationship from slide
    slide_rels = pkg.get_part_rels(slide_path)
    for _r_id, rel_info in slide_rels._rels.items():
        if rel_info.rel_type == REL_TYPE.NOTES_SLIDE:
            # Found notes slide
            if rel_info.target.startswith(".."):
                notes_path = f"ppt/notesSlides/{rel_info.target.split('/')[-1]}"
            else:
                notes_path = rel_info.target.lstrip("/")

            notes_xml = pkg.get_part(notes_path)
            if notes_xml:
                return NotesSlidePart.from_xml(notes_xml)

    return None


def create_notes_slide(pkg: Package, slide_number: int, text: str = "") -> NotesSlidePart:
    """Create a notes slide for a given slide, or update existing."""
    from .presentation import get_presentation_part

    pres_part = get_presentation_part(pkg)
    if pres_part is None:
        raise ValueError("Package has no presentation part")

    slide_refs = pres_part.get_slide_refs()
    if slide_number < 1 or slide_number > len(slide_refs):
        raise ValueError(f"Slide {slide_number} not found")

    ref = slide_refs[slide_number - 1]

    # Get slide path
    pres_rels = pkg.get_part_rels("ppt/presentation.xml")
    rel = pres_rels.get(ref.r_id)
    if rel is None:
        raise ValueError("Slide relationship not found")

    if rel.target.startswith("/"):
        slide_path = rel.target.lstrip("/")
    else:
        slide_path = f"ppt/{rel.target}"

    # Check if notes slide already exists
    slide_rels = pkg.get_part_rels(slide_path)
    existing_notes_path = None

    for _r_id, rel_info in slide_rels._rels.items():
        if rel_info.rel_type == REL_TYPE.NOTES_SLIDE:
            if rel_info.target.startswith(".."):
                existing_notes_path = f"ppt/notesSlides/{rel_info.target.split('/')[-1]}"
            else:
                existing_notes_path = rel_info.target.lstrip("/")
            break

    if existing_notes_path:
        # Update existing notes
        notes_xml = pkg.get_part(existing_notes_path)
        if notes_xml:
            notes_part = NotesSlidePart.from_xml(notes_xml)
            notes_part.set_text(text)
            pkg.set_part(existing_notes_path, notes_part.to_xml(), CONTENT_TYPE.NOTES_SLIDE)
            return notes_part

    # Create new notes slide
    notes_part = NotesSlidePart.new()
    if text:
        notes_part.set_text(text)

    # Determine notes slide number
    existing_notes = [
        name for name, _ in pkg.iter_parts()
        if name.startswith("ppt/notesSlides/notesSlide")
    ]
    notes_num = len(existing_notes) + 1
    notes_path = f"ppt/notesSlides/notesSlide{notes_num}.xml"

    # Add notes slide to package
    pkg.set_part(notes_path, notes_part.to_xml(), CONTENT_TYPE.NOTES_SLIDE)

    # Create notes slide relationships
    notes_rels = pkg.get_part_rels(notes_path)

    # Link to parent slide
    notes_rels.add(
        rel_type=REL_TYPE.SLIDE,
        target=f"../slides/slide{slide_number}.xml",
        r_id="rId1",
    )

    # Link to notes master (if exists)
    notes_master_path = "ppt/notesMasters/notesMaster1.xml"
    if pkg.get_part(notes_master_path):
        notes_rels.add(
            rel_type=REL_TYPE.NOTES_MASTER,
            target="../notesMasters/notesMaster1.xml",
            r_id="rId2",
        )

    pkg.set_part_rels(notes_path, notes_rels)

    # Add relationship from slide to notes
    slide_rels.add(
        rel_type=REL_TYPE.NOTES_SLIDE,
        target=f"../notesSlides/notesSlide{notes_num}.xml",
    )
    pkg.set_part_rels(slide_path, slide_rels)

    return notes_part
