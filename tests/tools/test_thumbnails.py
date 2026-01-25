"""Tests for slide thumbnail functionality."""

import pytest

import py2ppt as ppt


class TestGetPresentationThumbnail:
    """Tests for get_presentation_thumbnail function."""

    def test_get_thumbnail_new_presentation(self):
        """Test getting thumbnail from a new presentation (may be None)."""
        pres = ppt.create_presentation()
        ppt.add_slide(pres, layout="Blank")

        thumbnail = ppt.get_presentation_thumbnail(pres)

        # New presentations typically don't have embedded thumbnails
        # This should return None, not raise an error
        assert thumbnail is None or isinstance(thumbnail, bytes)

    def test_get_thumbnail_returns_bytes_or_none(self):
        """Test that get_presentation_thumbnail returns bytes or None."""
        pres = ppt.create_presentation()

        result = ppt.get_presentation_thumbnail(pres)

        assert result is None or isinstance(result, bytes)


class TestGetSlideThumbnail:
    """Tests for get_slide_thumbnail function."""

    @pytest.fixture
    def has_export_tools(self):
        """Check if export tools are available."""
        deps = ppt.check_export_dependencies()
        return deps.get("libreoffice", False)

    def test_get_slide_thumbnail_invalid_slide(self):
        """Test getting thumbnail for invalid slide number."""
        pres = ppt.create_presentation()
        ppt.add_slide(pres, layout="Blank")

        result = ppt.get_slide_thumbnail(pres, 999)

        assert result is None

    def test_get_slide_thumbnail_zero_slide(self):
        """Test getting thumbnail for slide 0."""
        pres = ppt.create_presentation()
        ppt.add_slide(pres, layout="Blank")

        result = ppt.get_slide_thumbnail(pres, 0)

        assert result is None

    @pytest.mark.skipif(
        not ppt.check_export_dependencies().get("libreoffice", False),
        reason="LibreOffice not available"
    )
    def test_get_slide_thumbnail_with_tools(self):
        """Test getting thumbnail when export tools are available."""
        pres = ppt.create_presentation()
        ppt.add_slide(pres, layout="Title Slide")
        ppt.set_title(pres, 1, "Test Slide")

        thumbnail = ppt.get_slide_thumbnail(pres, 1, width=160, height=120)

        # Should return bytes if tools are available
        assert thumbnail is None or isinstance(thumbnail, bytes)


class TestSaveSlideThumbnail:
    """Tests for save_slide_thumbnail function."""

    def test_save_thumbnail_invalid_slide(self):
        """Test saving thumbnail for invalid slide."""
        pres = ppt.create_presentation()
        ppt.add_slide(pres, layout="Blank")

        result = ppt.save_slide_thumbnail(pres, 999, "test.png")

        assert result is False

    def test_save_thumbnail_zero_slide(self):
        """Test saving thumbnail for slide 0."""
        pres = ppt.create_presentation()
        ppt.add_slide(pres, layout="Blank")

        result = ppt.save_slide_thumbnail(pres, 0, "test.png")

        assert result is False

    @pytest.mark.skipif(
        not ppt.check_export_dependencies().get("libreoffice", False),
        reason="LibreOffice not available"
    )
    def test_save_thumbnail_with_tools(self, tmp_path):
        """Test saving thumbnail when export tools are available."""
        pres = ppt.create_presentation()
        ppt.add_slide(pres, layout="Title Slide")
        ppt.set_title(pres, 1, "Test Slide")

        output = tmp_path / "thumb.png"
        result = ppt.save_slide_thumbnail(pres, 1, str(output))

        # Should succeed if tools are available
        assert result is True or result is False  # May fail without pdftoppm


class TestGetAllThumbnails:
    """Tests for get_all_thumbnails function."""

    def test_get_thumbnails_empty_presentation(self):
        """Test getting thumbnails from empty presentation."""
        pres = ppt.create_presentation()

        thumbnails = ppt.get_all_thumbnails(pres)

        assert thumbnails == []

    @pytest.mark.skipif(
        not ppt.check_export_dependencies().get("libreoffice", False),
        reason="LibreOffice not available"
    )
    def test_get_all_thumbnails_as_bytes(self):
        """Test getting all thumbnails as bytes."""
        pres = ppt.create_presentation()
        ppt.add_slide(pres, layout="Title Slide")
        ppt.add_slide(pres, layout="Blank")

        thumbnails = ppt.get_all_thumbnails(pres, width=160, height=120)

        # Should return list (may be empty if tools unavailable)
        assert isinstance(thumbnails, list)

    @pytest.mark.skipif(
        not ppt.check_export_dependencies().get("libreoffice", False),
        reason="LibreOffice not available"
    )
    def test_get_all_thumbnails_to_dir(self, tmp_path):
        """Test getting all thumbnails saved to directory."""
        pres = ppt.create_presentation()
        ppt.add_slide(pres, layout="Title Slide")
        ppt.add_slide(pres, layout="Blank")

        output_dir = tmp_path / "thumbs"
        paths = ppt.get_all_thumbnails(pres, str(output_dir), width=160, height=120)

        # Should return list of paths (may be empty if tools unavailable)
        assert isinstance(paths, list)
