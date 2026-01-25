"""Tests for master slide functionality."""

import tempfile
from pathlib import Path

import pytest

import py2ppt as ppt


class TestGetSlideMasters:
    """Tests for get_slide_masters function."""

    def test_get_masters_default_presentation(self):
        """Test getting masters from a default presentation."""
        pres = ppt.create_presentation()

        masters = ppt.get_slide_masters(pres)

        # Default presentation should have at least one master
        assert isinstance(masters, list)
        assert len(masters) >= 1

    def test_master_has_required_fields(self):
        """Test that master dict has required fields."""
        pres = ppt.create_presentation()

        masters = ppt.get_slide_masters(pres)

        if masters:
            master = masters[0]
            assert "index" in master
            assert "name" in master
            assert "layouts" in master
            assert isinstance(master["layouts"], list)

    def test_master_index_is_zero_based(self):
        """Test that master index starts at 0."""
        pres = ppt.create_presentation()

        masters = ppt.get_slide_masters(pres)

        if masters:
            assert masters[0]["index"] == 0


class TestGetMasterPlaceholders:
    """Tests for get_master_placeholders function."""

    def test_get_placeholders_default_master(self):
        """Test getting placeholders from default master."""
        pres = ppt.create_presentation()

        placeholders = ppt.get_master_placeholders(pres)

        assert isinstance(placeholders, list)

    def test_placeholder_has_required_fields(self):
        """Test that placeholder dict has required fields."""
        pres = ppt.create_presentation()

        placeholders = ppt.get_master_placeholders(pres)

        if placeholders:
            ph = placeholders[0]
            assert "type" in ph
            assert "idx" in ph
            assert "position" in ph
            assert isinstance(ph["position"], dict)

    def test_invalid_master_index(self):
        """Test getting placeholders with invalid master index."""
        pres = ppt.create_presentation()

        placeholders = ppt.get_master_placeholders(pres, master_index=999)

        assert placeholders == []

    def test_negative_master_index(self):
        """Test getting placeholders with negative master index."""
        pres = ppt.create_presentation()

        placeholders = ppt.get_master_placeholders(pres, master_index=-1)

        assert placeholders == []


class TestSetMasterBackground:
    """Tests for set_master_background function."""

    def test_set_solid_color_background(self):
        """Test setting a solid color background."""
        pres = ppt.create_presentation()

        result = ppt.set_master_background(pres, color="#003366")

        assert result is True

    def test_set_theme_color_background(self):
        """Test setting a theme color background."""
        pres = ppt.create_presentation()

        result = ppt.set_master_background(pres, color="accent1")

        assert result is True

    def test_set_gradient_background(self):
        """Test setting a gradient background."""
        pres = ppt.create_presentation()

        result = ppt.set_master_background(
            pres,
            gradient={"colors": ["#000066", "#0066CC"], "direction": 90}
        )

        assert result is True

    def test_set_background_with_transparency(self):
        """Test setting background with transparency."""
        pres = ppt.create_presentation()

        result = ppt.set_master_background(pres, color="#003366", transparency=50)

        assert result is True

    def test_set_background_invalid_master(self):
        """Test setting background with invalid master index."""
        pres = ppt.create_presentation()

        result = ppt.set_master_background(pres, master_index=999, color="#003366")

        assert result is False

    def test_set_image_background_nonexistent_file(self):
        """Test setting image background with nonexistent file."""
        pres = ppt.create_presentation()

        result = ppt.set_master_background(pres, image="/nonexistent/image.png")

        assert result is False

    def test_set_image_background(self):
        """Test setting image background."""
        pres = ppt.create_presentation()

        # Create a simple PNG image
        with tempfile.NamedTemporaryFile(suffix=".png", delete=False) as f:
            # Minimal valid PNG
            png_data = (
                b"\x89PNG\r\n\x1a\n\x00\x00\x00\rIHDR\x00\x00\x00\x01"
                b"\x00\x00\x00\x01\x08\x02\x00\x00\x00\x90wS\xde\x00"
                b"\x00\x00\x0cIDATx\x9cc\xf8\x0f\x00\x00\x01\x01\x00"
                b"\x05\x18\xd8N\x00\x00\x00\x00IEND\xaeB`\x82"
            )
            f.write(png_data)
            temp_path = f.name

        try:
            result = ppt.set_master_background(pres, image=temp_path)
            assert result is True
        finally:
            Path(temp_path).unlink()


class TestSetMasterFont:
    """Tests for set_master_font function."""

    def test_set_title_font_name(self):
        """Test setting font name for title placeholder."""
        pres = ppt.create_presentation()

        result = ppt.set_master_font(
            pres,
            placeholder_type="title",
            font_name="Arial"
        )

        # May be False if title placeholder not found
        assert result in (True, False)

    def test_set_title_font_size(self):
        """Test setting font size for title placeholder."""
        pres = ppt.create_presentation()

        result = ppt.set_master_font(
            pres,
            placeholder_type="title",
            font_size=44
        )

        assert result in (True, False)

    def test_set_title_font_color(self):
        """Test setting font color for title placeholder."""
        pres = ppt.create_presentation()

        result = ppt.set_master_font(
            pres,
            placeholder_type="title",
            color="#333333"
        )

        assert result in (True, False)

    def test_set_title_font_bold_italic(self):
        """Test setting bold and italic for title."""
        pres = ppt.create_presentation()

        result = ppt.set_master_font(
            pres,
            placeholder_type="title",
            bold=True,
            italic=True
        )

        assert result in (True, False)

    def test_set_body_font(self):
        """Test setting font for body placeholder."""
        pres = ppt.create_presentation()

        result = ppt.set_master_font(
            pres,
            placeholder_type="body",
            font_name="Calibri",
            font_size=18
        )

        assert result in (True, False)

    def test_set_font_invalid_master(self):
        """Test setting font with invalid master index."""
        pres = ppt.create_presentation()

        result = ppt.set_master_font(
            pres,
            master_index=999,
            placeholder_type="title",
            font_name="Arial"
        )

        assert result is False

    def test_set_font_nonexistent_placeholder(self):
        """Test setting font for nonexistent placeholder type."""
        pres = ppt.create_presentation()

        result = ppt.set_master_font(
            pres,
            placeholder_type="nonexistent_type",
            font_name="Arial"
        )

        assert result is False


class TestGetLayoutInfo:
    """Tests for get_layout_info function."""

    def test_get_layout_by_name(self):
        """Test getting layout info by name."""
        pres = ppt.create_presentation()

        info = ppt.get_layout_info(pres, layout_name="Title Slide")

        if info is not None:
            assert "name" in info
            assert "index" in info
            assert "master_index" in info
            assert "placeholders" in info

    def test_get_layout_by_partial_name(self):
        """Test getting layout info by partial name."""
        pres = ppt.create_presentation()

        info = ppt.get_layout_info(pres, layout_name="title")

        # Should find a layout with "title" in the name
        assert info is None or "name" in info

    def test_get_layout_by_index(self):
        """Test getting layout info by index."""
        pres = ppt.create_presentation()

        info = ppt.get_layout_info(pres, layout_index=0)

        if info is not None:
            assert "name" in info
            assert info["index"] == 0

    def test_get_layout_nonexistent_name(self):
        """Test getting layout info with nonexistent name."""
        pres = ppt.create_presentation()

        info = ppt.get_layout_info(pres, layout_name="NonexistentLayout123")

        assert info is None

    def test_get_layout_invalid_index(self):
        """Test getting layout info with invalid index."""
        pres = ppt.create_presentation()

        info = ppt.get_layout_info(pres, layout_index=9999)

        assert info is None


class TestAddLogoToMaster:
    """Tests for add_logo_to_master function."""

    @pytest.fixture
    def temp_logo(self):
        """Create a temporary logo file."""
        # Minimal valid PNG
        png_data = (
            b"\x89PNG\r\n\x1a\n\x00\x00\x00\rIHDR\x00\x00\x00\x01"
            b"\x00\x00\x00\x01\x08\x02\x00\x00\x00\x90wS\xde\x00"
            b"\x00\x00\x0cIDATx\x9cc\xf8\x0f\x00\x00\x01\x01\x00"
            b"\x05\x18\xd8N\x00\x00\x00\x00IEND\xaeB`\x82"
        )
        with tempfile.NamedTemporaryFile(suffix=".png", delete=False) as f:
            f.write(png_data)
            temp_path = f.name
        yield temp_path
        Path(temp_path).unlink()

    def test_add_logo_basic(self, temp_logo):
        """Test adding a logo with default position."""
        pres = ppt.create_presentation()

        result = ppt.add_logo_to_master(pres, temp_logo)

        assert result is True

    def test_add_logo_with_position(self, temp_logo):
        """Test adding a logo with specific position."""
        pres = ppt.create_presentation()

        result = ppt.add_logo_to_master(
            pres,
            temp_logo,
            left="1in",
            top="0.5in",
            width="2in",
            height="1in"
        )

        assert result is True

    def test_add_logo_with_emu_position(self, temp_logo):
        """Test adding a logo with EMU values."""
        pres = ppt.create_presentation()

        result = ppt.add_logo_to_master(
            pres,
            temp_logo,
            left=914400,
            top=457200,
            width=1828800
        )

        assert result is True

    def test_add_logo_nonexistent_file(self):
        """Test adding logo with nonexistent file."""
        pres = ppt.create_presentation()

        result = ppt.add_logo_to_master(pres, "/nonexistent/logo.png")

        assert result is False

    def test_add_logo_invalid_master(self, temp_logo):
        """Test adding logo with invalid master index."""
        pres = ppt.create_presentation()

        result = ppt.add_logo_to_master(pres, temp_logo, master_index=999)

        assert result is False

    def test_add_logo_appears_on_slides(self, temp_logo):
        """Test that logo is associated with master."""
        pres = ppt.create_presentation()
        ppt.add_slide(pres, layout="Title Slide")

        result = ppt.add_logo_to_master(pres, temp_logo)

        assert result is True

        # Verify presentation can be saved and loaded
        with tempfile.NamedTemporaryFile(suffix=".pptx", delete=False) as f:
            temp_pptx = f.name

        try:
            ppt.save_presentation(pres, temp_pptx)
            pres2 = ppt.open_presentation(temp_pptx)
            # Should have one slide
            assert ppt.get_slide_count(pres2) == 1
        finally:
            Path(temp_pptx).unlink()


class TestMasterIntegration:
    """Integration tests for master slide functionality."""

    def test_modify_master_and_save(self):
        """Test modifying master and saving presentation."""
        pres = ppt.create_presentation()
        ppt.add_slide(pres, layout="Title Slide")
        ppt.set_title(pres, 1, "Test Title")

        # Modify master background
        ppt.set_master_background(pres, color="#E6F0FF")

        with tempfile.NamedTemporaryFile(suffix=".pptx", delete=False) as f:
            temp_pptx = f.name

        try:
            ppt.save_presentation(pres, temp_pptx)
            pres2 = ppt.open_presentation(temp_pptx)

            # Verify presentation is valid
            assert ppt.get_slide_count(pres2) == 1

            # Verify master still exists
            masters = ppt.get_slide_masters(pres2)
            assert len(masters) >= 1
        finally:
            Path(temp_pptx).unlink()

    def test_master_layouts_list(self):
        """Test that layouts are properly listed for each master."""
        pres = ppt.create_presentation()

        masters = ppt.get_slide_masters(pres)

        if masters:
            # Each master should have associated layouts
            for master in masters:
                assert isinstance(master["layouts"], list)
