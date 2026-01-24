"""PowerPoint package (.pptx) handling.

A .pptx file is a ZIP archive containing XML files and other resources.
This module handles reading/writing the package structure.
"""

from __future__ import annotations

import io
import os
import zipfile
from dataclasses import dataclass, field
from pathlib import Path
from typing import BinaryIO, Dict, Iterator, List, Optional, Union
from xml.etree import ElementTree as ET

from lxml import etree

from .ns import CONTENT_TYPE, REL_TYPE, nsmap, qn


@dataclass
class Relationship:
    """Represents a relationship in a .rels file."""

    r_id: str
    rel_type: str
    target: str
    target_mode: str = "Internal"  # "Internal" or "External"

    def to_element(self) -> etree._Element:
        """Convert to XML element."""
        attrib = {
            "Id": self.r_id,
            "Type": self.rel_type,
            "Target": self.target,
        }
        if self.target_mode == "External":
            attrib["TargetMode"] = "External"

        return etree.Element(
            qn("pr:Relationship"),
            attrib=attrib,
            nsmap={"": nsmap["pr"]},
        )


@dataclass
class ContentTypeOverride:
    """Content type override entry."""

    part_name: str
    content_type: str


@dataclass
class ContentTypeDefault:
    """Content type default entry."""

    extension: str
    content_type: str


class RelationshipCollection:
    """Manages relationships for a package part."""

    def __init__(self) -> None:
        self._rels: Dict[str, Relationship] = {}
        self._next_id: int = 1

    def __iter__(self) -> Iterator[Relationship]:
        return iter(self._rels.values())

    def __len__(self) -> int:
        return len(self._rels)

    def get(self, r_id: str) -> Optional[Relationship]:
        """Get relationship by ID."""
        return self._rels.get(r_id)

    def find_by_type(self, rel_type: str) -> List[Relationship]:
        """Find all relationships of a given type."""
        return [r for r in self._rels.values() if r.rel_type == rel_type]

    def add(
        self,
        rel_type: str,
        target: str,
        r_id: Optional[str] = None,
        target_mode: str = "Internal",
    ) -> str:
        """Add a relationship and return its ID."""
        if r_id is None:
            r_id = f"rId{self._next_id}"
            self._next_id += 1

        self._rels[r_id] = Relationship(
            r_id=r_id,
            rel_type=rel_type,
            target=target,
            target_mode=target_mode,
        )

        # Update next_id if needed
        if r_id.startswith("rId"):
            try:
                num = int(r_id[3:])
                if num >= self._next_id:
                    self._next_id = num + 1
            except ValueError:
                pass

        return r_id

    def remove(self, r_id: str) -> None:
        """Remove a relationship by ID."""
        if r_id in self._rels:
            del self._rels[r_id]

    def to_xml(self) -> bytes:
        """Serialize relationships to XML."""
        root = etree.Element(
            qn("pr:Relationships"),
            nsmap={None: nsmap["pr"]},
        )

        for rel in sorted(self._rels.values(), key=lambda r: r.r_id):
            attrib = {
                "Id": rel.r_id,
                "Type": rel.rel_type,
                "Target": rel.target,
            }
            if rel.target_mode == "External":
                attrib["TargetMode"] = "External"

            etree.SubElement(root, qn("pr:Relationship"), attrib=attrib)

        return etree.tostring(
            root,
            xml_declaration=True,
            encoding="UTF-8",
            standalone=True,
        )

    @classmethod
    def from_xml(cls, xml_bytes: bytes) -> "RelationshipCollection":
        """Parse relationships from XML."""
        coll = cls()
        root = etree.fromstring(xml_bytes)

        for rel_elem in root.findall(qn("pr:Relationship")):
            r_id = rel_elem.get("Id", "")
            rel_type = rel_elem.get("Type", "")
            target = rel_elem.get("Target", "")
            target_mode = rel_elem.get("TargetMode", "Internal")

            coll.add(
                rel_type=rel_type,
                target=target,
                r_id=r_id,
                target_mode=target_mode,
            )

        return coll


class ContentTypes:
    """Manages [Content_Types].xml."""

    def __init__(self) -> None:
        self.defaults: Dict[str, str] = {
            "rels": CONTENT_TYPE.RELS,
            "xml": CONTENT_TYPE.XML,
        }
        self.overrides: Dict[str, str] = {}

    def add_default(self, extension: str, content_type: str) -> None:
        """Add a default content type for an extension."""
        self.defaults[extension] = content_type

    def add_override(self, part_name: str, content_type: str) -> None:
        """Add a content type override for a specific part."""
        if not part_name.startswith("/"):
            part_name = "/" + part_name
        self.overrides[part_name] = content_type

    def remove_override(self, part_name: str) -> None:
        """Remove a content type override."""
        if not part_name.startswith("/"):
            part_name = "/" + part_name
        if part_name in self.overrides:
            del self.overrides[part_name]

    def get_content_type(self, part_name: str) -> Optional[str]:
        """Get content type for a part."""
        if not part_name.startswith("/"):
            part_name = "/" + part_name

        if part_name in self.overrides:
            return self.overrides[part_name]

        ext = part_name.rsplit(".", 1)[-1] if "." in part_name else ""
        return self.defaults.get(ext)

    def to_xml(self) -> bytes:
        """Serialize to XML."""
        root = etree.Element(
            qn("ct:Types"),
            nsmap={None: nsmap["ct"]},
        )

        for ext, ct in sorted(self.defaults.items()):
            etree.SubElement(
                root,
                qn("ct:Default"),
                attrib={"Extension": ext, "ContentType": ct},
            )

        for part, ct in sorted(self.overrides.items()):
            etree.SubElement(
                root,
                qn("ct:Override"),
                attrib={"PartName": part, "ContentType": ct},
            )

        return etree.tostring(
            root,
            xml_declaration=True,
            encoding="UTF-8",
            standalone=True,
        )

    @classmethod
    def from_xml(cls, xml_bytes: bytes) -> "ContentTypes":
        """Parse from XML."""
        ct = cls()
        ct.defaults.clear()

        root = etree.fromstring(xml_bytes)

        for default in root.findall(qn("ct:Default")):
            ext = default.get("Extension", "")
            content_type = default.get("ContentType", "")
            if ext and content_type:
                ct.defaults[ext] = content_type

        for override in root.findall(qn("ct:Override")):
            part = override.get("PartName", "")
            content_type = override.get("ContentType", "")
            if part and content_type:
                ct.overrides[part] = content_type

        return ct


class Package:
    """Represents a PowerPoint package (.pptx file).

    A package is a ZIP file containing:
    - [Content_Types].xml - content type manifest
    - _rels/.rels - package relationships
    - ppt/ - presentation parts (slides, layouts, masters, theme)
    - docProps/ - document properties
    """

    def __init__(self) -> None:
        self.content_types = ContentTypes()
        self.package_rels = RelationshipCollection()
        self._parts: Dict[str, bytes] = {}
        self._part_rels: Dict[str, RelationshipCollection] = {}

    def get_part(self, part_name: str) -> Optional[bytes]:
        """Get raw bytes for a part."""
        part_name = part_name.lstrip("/")
        return self._parts.get(part_name)

    def set_part(
        self,
        part_name: str,
        content: bytes,
        content_type: Optional[str] = None,
    ) -> None:
        """Set content for a part."""
        part_name = part_name.lstrip("/")
        self._parts[part_name] = content

        if content_type:
            self.content_types.add_override("/" + part_name, content_type)

    def remove_part(self, part_name: str) -> None:
        """Remove a part from the package."""
        part_name = part_name.lstrip("/")
        if part_name in self._parts:
            del self._parts[part_name]
        self.content_types.remove_override("/" + part_name)

        # Remove relationships
        rels_path = self._rels_path_for(part_name)
        if rels_path in self._parts:
            del self._parts[rels_path]
        if part_name in self._part_rels:
            del self._part_rels[part_name]

    def get_part_rels(self, part_name: str) -> RelationshipCollection:
        """Get relationships for a part."""
        part_name = part_name.lstrip("/")

        if part_name not in self._part_rels:
            # Try loading from package
            rels_path = self._rels_path_for(part_name)
            if rels_path in self._parts:
                self._part_rels[part_name] = RelationshipCollection.from_xml(
                    self._parts[rels_path]
                )
            else:
                self._part_rels[part_name] = RelationshipCollection()

        return self._part_rels[part_name]

    def set_part_rels(
        self, part_name: str, rels: RelationshipCollection
    ) -> None:
        """Set relationships for a part."""
        part_name = part_name.lstrip("/")
        self._part_rels[part_name] = rels

    def _rels_path_for(self, part_name: str) -> str:
        """Get the .rels path for a part."""
        part_name = part_name.lstrip("/")
        if "/" in part_name:
            dir_part, file_part = part_name.rsplit("/", 1)
            return f"{dir_part}/_rels/{file_part}.rels"
        return f"_rels/{part_name}.rels"

    def iter_parts(self) -> Iterator[tuple[str, bytes]]:
        """Iterate over all parts."""
        for name, content in self._parts.items():
            yield name, content

    @classmethod
    def open(cls, path_or_file: Union[str, Path, BinaryIO]) -> "Package":
        """Open a package from file or file-like object."""
        pkg = cls()

        if isinstance(path_or_file, (str, Path)):
            zf = zipfile.ZipFile(path_or_file, "r")
            should_close = True
        else:
            zf = zipfile.ZipFile(path_or_file, "r")
            should_close = True

        try:
            # Read all parts
            for name in zf.namelist():
                pkg._parts[name] = zf.read(name)

            # Parse content types
            if "[Content_Types].xml" in pkg._parts:
                pkg.content_types = ContentTypes.from_xml(
                    pkg._parts["[Content_Types].xml"]
                )
                del pkg._parts["[Content_Types].xml"]

            # Parse package rels
            if "_rels/.rels" in pkg._parts:
                pkg.package_rels = RelationshipCollection.from_xml(
                    pkg._parts["_rels/.rels"]
                )
                del pkg._parts["_rels/.rels"]

            # Parse part rels (lazy - done in get_part_rels)

        finally:
            if should_close:
                zf.close()

        return pkg

    def save(self, path_or_file: Union[str, Path, BinaryIO]) -> None:
        """Save package to file or file-like object."""
        if isinstance(path_or_file, (str, Path)):
            zf = zipfile.ZipFile(
                path_or_file,
                "w",
                compression=zipfile.ZIP_DEFLATED,
            )
            should_close = True
        else:
            zf = zipfile.ZipFile(
                path_or_file,
                "w",
                compression=zipfile.ZIP_DEFLATED,
            )
            should_close = True

        try:
            # Write content types
            zf.writestr("[Content_Types].xml", self.content_types.to_xml())

            # Write package rels
            if len(self.package_rels) > 0:
                zf.writestr("_rels/.rels", self.package_rels.to_xml())

            # Write part rels
            for part_name, rels in self._part_rels.items():
                if len(rels) > 0:
                    rels_path = self._rels_path_for(part_name)
                    zf.writestr(rels_path, rels.to_xml())

            # Write parts
            for name, content in self._parts.items():
                # Skip rels files (handled above)
                if "/_rels/" in name or name.startswith("_rels/"):
                    if name.endswith(".rels"):
                        continue
                zf.writestr(name, content)

        finally:
            if should_close:
                zf.close()

    def to_bytes(self) -> bytes:
        """Serialize package to bytes."""
        buf = io.BytesIO()
        self.save(buf)
        return buf.getvalue()

    @classmethod
    def from_bytes(cls, data: bytes) -> "Package":
        """Load package from bytes."""
        return cls.open(io.BytesIO(data))


def create_blank_package() -> Package:
    """Create a minimal blank PowerPoint package."""
    pkg = Package()

    # Add default content types
    pkg.content_types.add_default("jpeg", CONTENT_TYPE.JPEG)
    pkg.content_types.add_default("png", CONTENT_TYPE.PNG)
    pkg.content_types.add_default("gif", CONTENT_TYPE.GIF)
    pkg.content_types.add_default("emf", CONTENT_TYPE.EMF)

    # Add package rels
    pkg.package_rels.add(
        rel_type=REL_TYPE.OFFICE_DOCUMENT,
        target="ppt/presentation.xml",
        r_id="rId1",
    )
    pkg.package_rels.add(
        rel_type=REL_TYPE.CORE_PROPS,
        target="docProps/core.xml",
        r_id="rId2",
    )
    pkg.package_rels.add(
        rel_type=REL_TYPE.EXT_PROPS,
        target="docProps/app.xml",
        r_id="rId3",
    )

    # Core properties
    core_xml = b'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties"
    xmlns:dc="http://purl.org/dc/elements/1.1/"
    xmlns:dcterms="http://purl.org/dc/terms/"
    xmlns:dcmitype="http://purl.org/dc/dcmitype/"
    xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
    <dc:creator>py2ppt</dc:creator>
    <cp:lastModifiedBy>py2ppt</cp:lastModifiedBy>
</cp:coreProperties>'''
    pkg.set_part("docProps/core.xml", core_xml, CONTENT_TYPE.CORE_PROPS)

    # Extended properties
    app_xml = b'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"
    xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">
    <Application>py2ppt</Application>
    <Slides>0</Slides>
</Properties>'''
    pkg.set_part("docProps/app.xml", app_xml, CONTENT_TYPE.EXT_PROPS)

    return pkg
