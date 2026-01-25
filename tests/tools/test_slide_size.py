"""Tests for slide size functionality."""

import pytest

import py2ppt as ppt


class TestGetSlideSize:
    """Tests for get_slide_size function."""

    def test_get_slide_size(self):
        """Test getting slide dimensions."""
        pres = ppt.create_presentation()

        size = ppt.get_slide_size(pres)

        assert "width" in size
        assert "height" in size
        assert "width_inches" in size
        assert "height_inches" in size
        assert "aspect_ratio" in size

    def test_get_slide_size_values(self):
        """Test slide size has reasonable values."""
        pres = ppt.create_presentation()

        size = ppt.get_slide_size(pres)

        # EMU values should be positive
        assert size["width"] > 0
        assert size["height"] > 0

        # Inches should be reasonable (typical slides are 10-13" wide)
        assert 5 <= size["width_inches"] <= 20
        assert 5 <= size["height_inches"] <= 15


class TestSetSlideSize:
    """Tests for set_slide_size function."""

    def test_set_slide_size_inches(self):
        """Test setting slide size in inches."""
        pres = ppt.create_presentation()

        ppt.set_slide_size(pres, "10in", "7.5in")
        size = ppt.get_slide_size(pres)

        assert abs(size["width_inches"] - 10) < 0.1
        assert abs(size["height_inches"] - 7.5) < 0.1

    def test_set_slide_size_cm(self):
        """Test setting slide size in centimeters."""
        pres = ppt.create_presentation()

        # 25.4 cm = 10 inches
        ppt.set_slide_size(pres, "25.4cm", "19.05cm")
        size = ppt.get_slide_size(pres)

        assert abs(size["width_inches"] - 10) < 0.2


class TestSetSlideSizePreset:
    """Tests for set_slide_size_preset function."""

    def test_set_widescreen(self):
        """Test setting widescreen (16:9) preset."""
        pres = ppt.create_presentation()

        ppt.set_slide_size_preset(pres, "widescreen")
        size = ppt.get_slide_size(pres)

        assert size["aspect_ratio"] == "16:9"

    def test_set_standard(self):
        """Test setting standard (4:3) preset."""
        pres = ppt.create_presentation()

        ppt.set_slide_size_preset(pres, "standard")
        size = ppt.get_slide_size(pres)

        assert size["aspect_ratio"] == "4:3"

    def test_invalid_preset(self):
        """Test invalid preset raises error."""
        pres = ppt.create_presentation()

        with pytest.raises(ValueError):
            ppt.set_slide_size_preset(pres, "invalid_preset_name")
