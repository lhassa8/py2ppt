"""Tests for section functionality."""

import py2ppt as ppt


class TestAddSection:
    """Tests for add_section function."""

    def test_add_section(self):
        """Test adding a section."""
        pres = ppt.create_presentation()
        ppt.add_slide(pres, layout="Title and Content")

        section_id = ppt.add_section(pres, "Introduction")

        assert section_id >= 0

    def test_add_section_before_slide(self):
        """Test adding a section before a specific slide."""
        pres = ppt.create_presentation()
        ppt.add_slide(pres, layout="Title and Content")
        ppt.add_slide(pres, layout="Title and Content")

        section_id = ppt.add_section(pres, "Chapter 2", before_slide=2)

        assert section_id >= 0


class TestGetSections:
    """Tests for get_sections function."""

    def test_get_sections_empty(self):
        """Test getting sections from presentation without sections."""
        pres = ppt.create_presentation()
        ppt.add_slide(pres, layout="Title and Content")

        sections = ppt.get_sections(pres)

        assert isinstance(sections, list)

    def test_get_sections_with_sections(self):
        """Test getting sections after adding them."""
        pres = ppt.create_presentation()
        ppt.add_slide(pres, layout="Title and Content")
        ppt.add_section(pres, "Introduction")

        sections = ppt.get_sections(pres)

        assert len(sections) >= 1
        assert sections[0]["name"] == "Introduction"


class TestRenameSection:
    """Tests for rename_section function."""

    def test_rename_section(self):
        """Test renaming a section."""
        pres = ppt.create_presentation()
        ppt.add_slide(pres, layout="Title and Content")
        ppt.add_section(pres, "Introduction")

        result = ppt.rename_section(pres, 0, "Executive Summary")

        assert result is True
        sections = ppt.get_sections(pres)
        assert sections[0]["name"] == "Executive Summary"

    def test_rename_nonexistent_section(self):
        """Test renaming a section that doesn't exist."""
        pres = ppt.create_presentation()
        ppt.add_slide(pres, layout="Title and Content")

        result = ppt.rename_section(pres, 99, "New Name")

        assert result is False


class TestDeleteSection:
    """Tests for delete_section function."""

    def test_delete_section(self):
        """Test deleting a section."""
        pres = ppt.create_presentation()
        ppt.add_slide(pres, layout="Title and Content")
        ppt.add_section(pres, "Introduction")

        result = ppt.delete_section(pres, 0)

        assert result is True
        sections = ppt.get_sections(pres)
        assert len(sections) == 0

    def test_delete_nonexistent_section(self):
        """Test deleting a section that doesn't exist."""
        pres = ppt.create_presentation()
        ppt.add_slide(pres, layout="Title and Content")

        result = ppt.delete_section(pres, 99)

        assert result is False
