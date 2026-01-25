"""Tests for document properties functionality."""

import py2ppt as ppt


class TestSetDocumentProperty:
    """Tests for set_document_property function."""

    def test_set_title(self):
        """Test setting document title."""
        pres = ppt.create_presentation()

        ppt.set_document_property(pres, "title", "My Presentation")

        assert ppt.get_document_property(pres, "title") == "My Presentation"

    def test_set_author(self):
        """Test setting document author."""
        pres = ppt.create_presentation()

        ppt.set_document_property(pres, "author", "John Smith")

        assert ppt.get_document_property(pres, "author") == "John Smith"

    def test_set_subject(self):
        """Test setting document subject."""
        pres = ppt.create_presentation()

        ppt.set_document_property(pres, "subject", "Quarterly Report")

        assert ppt.get_document_property(pres, "subject") == "Quarterly Report"

    def test_set_keywords(self):
        """Test setting document keywords."""
        pres = ppt.create_presentation()

        ppt.set_document_property(pres, "keywords", "finance, quarterly, report")

        assert ppt.get_document_property(pres, "keywords") == "finance, quarterly, report"

    def test_set_multiple_properties(self):
        """Test setting multiple properties."""
        pres = ppt.create_presentation()

        ppt.set_document_property(pres, "title", "Annual Report")
        ppt.set_document_property(pres, "author", "Finance Team")
        ppt.set_document_property(pres, "category", "Financial")

        assert ppt.get_document_property(pres, "title") == "Annual Report"
        assert ppt.get_document_property(pres, "author") == "Finance Team"
        assert ppt.get_document_property(pres, "category") == "Financial"

    def test_update_existing_property(self):
        """Test updating an existing property."""
        pres = ppt.create_presentation()

        ppt.set_document_property(pres, "title", "Draft")
        ppt.set_document_property(pres, "title", "Final")

        assert ppt.get_document_property(pres, "title") == "Final"


class TestGetDocumentProperty:
    """Tests for get_document_property function."""

    def test_get_nonexistent_property(self):
        """Test getting property that doesn't exist."""
        pres = ppt.create_presentation()

        result = ppt.get_document_property(pres, "title")

        assert result is None

    def test_get_unknown_property(self):
        """Test getting unknown property name."""
        pres = ppt.create_presentation()

        result = ppt.get_document_property(pres, "unknown_property")

        assert result is None


class TestGetDocumentInfo:
    """Tests for get_document_info function."""

    def test_get_document_info(self):
        """Test getting all document info."""
        pres = ppt.create_presentation()

        ppt.set_document_property(pres, "title", "Test")
        ppt.set_document_property(pres, "author", "Tester")

        info = ppt.get_document_info(pres)

        assert isinstance(info, dict)
        assert info.get("title") == "Test"
        assert info.get("author") == "Tester"

    def test_get_document_info_empty(self):
        """Test getting info from new presentation."""
        pres = ppt.create_presentation()

        info = ppt.get_document_info(pres)

        assert isinstance(info, dict)


class TestSetCustomProperty:
    """Tests for set_custom_property function."""

    def test_set_string_property(self):
        """Test setting string custom property."""
        pres = ppt.create_presentation()

        ppt.set_custom_property(pres, "Project Code", "PRJ-2024-001")

        assert ppt.get_custom_property(pres, "Project Code") == "PRJ-2024-001"

    def test_set_integer_property(self):
        """Test setting integer custom property."""
        pres = ppt.create_presentation()

        ppt.set_custom_property(pres, "Version", 5)

        assert ppt.get_custom_property(pres, "Version") == 5

    def test_set_float_property(self):
        """Test setting float custom property."""
        pres = ppt.create_presentation()

        ppt.set_custom_property(pres, "Rating", 4.5)

        result = ppt.get_custom_property(pres, "Rating")
        assert abs(result - 4.5) < 0.001

    def test_set_bool_property_true(self):
        """Test setting boolean custom property (True)."""
        pres = ppt.create_presentation()

        ppt.set_custom_property(pres, "Approved", True)

        assert ppt.get_custom_property(pres, "Approved") is True

    def test_set_bool_property_false(self):
        """Test setting boolean custom property (False)."""
        pres = ppt.create_presentation()

        ppt.set_custom_property(pres, "Draft", False)

        assert ppt.get_custom_property(pres, "Draft") is False

    def test_update_custom_property(self):
        """Test updating existing custom property."""
        pres = ppt.create_presentation()

        ppt.set_custom_property(pres, "Status", "Draft")
        ppt.set_custom_property(pres, "Status", "Final")

        assert ppt.get_custom_property(pres, "Status") == "Final"


class TestGetCustomProperty:
    """Tests for get_custom_property function."""

    def test_get_nonexistent_property(self):
        """Test getting custom property that doesn't exist."""
        pres = ppt.create_presentation()

        result = ppt.get_custom_property(pres, "NonExistent")

        assert result is None


class TestGetCustomProperties:
    """Tests for get_custom_properties function."""

    def test_get_all_custom_properties(self):
        """Test getting all custom properties."""
        pres = ppt.create_presentation()

        ppt.set_custom_property(pres, "Project", "Alpha")
        ppt.set_custom_property(pres, "Version", 1)
        ppt.set_custom_property(pres, "Active", True)

        props = ppt.get_custom_properties(pres)

        assert isinstance(props, dict)
        assert props["Project"] == "Alpha"
        assert props["Version"] == 1
        assert props["Active"] is True

    def test_get_custom_properties_empty(self):
        """Test getting custom properties when none exist."""
        pres = ppt.create_presentation()

        props = ppt.get_custom_properties(pres)

        assert isinstance(props, dict)
        assert len(props) == 0


class TestRemoveCustomProperty:
    """Tests for remove_custom_property function."""

    def test_remove_custom_property(self):
        """Test removing a custom property."""
        pres = ppt.create_presentation()

        ppt.set_custom_property(pres, "ToRemove", "value")
        result = ppt.remove_custom_property(pres, "ToRemove")

        assert result is True
        assert ppt.get_custom_property(pres, "ToRemove") is None

    def test_remove_nonexistent_property(self):
        """Test removing property that doesn't exist."""
        pres = ppt.create_presentation()

        result = ppt.remove_custom_property(pres, "NonExistent")

        assert result is False

    def test_remove_preserves_other_properties(self):
        """Test that removing one property preserves others."""
        pres = ppt.create_presentation()

        ppt.set_custom_property(pres, "Keep", "value1")
        ppt.set_custom_property(pres, "Remove", "value2")

        ppt.remove_custom_property(pres, "Remove")

        assert ppt.get_custom_property(pres, "Keep") == "value1"
        assert ppt.get_custom_property(pres, "Remove") is None
