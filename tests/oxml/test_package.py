"""Tests for package handling."""

import io
import zipfile

from py2ppt.oxml.package import (
    ContentTypes,
    Package,
    RelationshipCollection,
    create_blank_package,
)


class TestRelationshipCollection:
    """Tests for RelationshipCollection."""

    def test_add_relationship(self):
        coll = RelationshipCollection()
        r_id = coll.add(
            rel_type="http://example.com/type",
            target="target.xml",
        )
        assert r_id == "rId1"

    def test_add_with_custom_id(self):
        coll = RelationshipCollection()
        r_id = coll.add(
            rel_type="http://example.com/type",
            target="target.xml",
            r_id="rId5",
        )
        assert r_id == "rId5"

    def test_get_relationship(self):
        coll = RelationshipCollection()
        coll.add(
            rel_type="http://example.com/type",
            target="target.xml",
            r_id="rId1",
        )
        rel = coll.get("rId1")
        assert rel is not None
        assert rel.target == "target.xml"

    def test_find_by_type(self):
        coll = RelationshipCollection()
        coll.add(rel_type="type1", target="a.xml")
        coll.add(rel_type="type2", target="b.xml")
        coll.add(rel_type="type1", target="c.xml")

        results = coll.find_by_type("type1")
        assert len(results) == 2

    def test_to_xml_and_from_xml(self):
        coll = RelationshipCollection()
        coll.add(rel_type="http://example.com/type", target="target.xml")

        xml = coll.to_xml()
        assert b"Relationship" in xml

        parsed = RelationshipCollection.from_xml(xml)
        assert len(parsed) == 1


class TestContentTypes:
    """Tests for ContentTypes."""

    def test_default_content_types(self):
        ct = ContentTypes()
        assert ct.defaults["rels"] is not None
        assert ct.defaults["xml"] is not None

    def test_add_override(self):
        ct = ContentTypes()
        ct.add_override("/ppt/slides/slide1.xml", "application/xml")
        assert ct.overrides["/ppt/slides/slide1.xml"] == "application/xml"

    def test_get_content_type_from_override(self):
        ct = ContentTypes()
        ct.add_override("/test.xml", "custom/type")
        assert ct.get_content_type("/test.xml") == "custom/type"

    def test_get_content_type_from_default(self):
        ct = ContentTypes()
        assert "xml" in ct.get_content_type("/test.xml")


class TestPackage:
    """Tests for Package."""

    def test_create_blank_package(self):
        pkg = create_blank_package()
        assert pkg is not None
        assert len(pkg.package_rels) > 0

    def test_set_and_get_part(self):
        pkg = Package()
        pkg.set_part("test.xml", b"<test/>")
        assert pkg.get_part("test.xml") == b"<test/>"

    def test_remove_part(self):
        pkg = Package()
        pkg.set_part("test.xml", b"<test/>")
        pkg.remove_part("test.xml")
        assert pkg.get_part("test.xml") is None

    def test_save_and_open(self):
        # Create a package
        pkg = create_blank_package()
        pkg.set_part("custom/test.xml", b"<test>Hello</test>")

        # Save to bytes
        data = pkg.to_bytes()

        # Verify it's a valid ZIP
        zf = zipfile.ZipFile(io.BytesIO(data))
        assert "[Content_Types].xml" in zf.namelist()
        zf.close()

        # Reopen
        pkg2 = Package.from_bytes(data)
        assert pkg2.get_part("custom/test.xml") == b"<test>Hello</test>"
