"""Tests for theme functionality."""


import py2ppt as ppt


class TestGetThemeInfo:
    """Tests for get_theme_info function."""

    def test_get_theme_info(self):
        """Test getting theme information."""
        pres = ppt.create_presentation()

        info = ppt.get_theme_info(pres)

        assert isinstance(info, dict)
        assert "colors" in info
        assert "fonts" in info
        assert "name" in info

    def test_theme_has_accent_colors(self):
        """Test that theme info includes accent colors."""
        pres = ppt.create_presentation()
        info = ppt.get_theme_info(pres)

        colors = info["colors"]
        # Should have standard theme color slots
        assert "dk1" in colors or "accent1" in colors


class TestSetThemeColor:
    """Tests for set_theme_color function."""

    def test_set_accent_color(self, tmp_path):
        """Test setting an accent color."""
        pres = ppt.create_presentation()
        ppt.add_slide(pres, layout="Title Slide")  # Ensure theme is loaded

        ppt.set_theme_color(pres, "accent1", "#FF0000")

        path = tmp_path / "custom_theme.pptx"
        ppt.save_presentation(pres, str(path))
        assert path.exists()

    def test_set_multiple_colors(self, tmp_path):
        """Test setting multiple theme colors."""
        pres = ppt.create_presentation()
        ppt.add_slide(pres, layout="Title Slide")  # Ensure theme is loaded

        ppt.set_theme_color(pres, "accent1", "#FF0000")
        ppt.set_theme_color(pres, "accent2", "#00FF00")
        ppt.set_theme_color(pres, "accent3", "#0000FF")

        path = tmp_path / "multi_color_theme.pptx"
        ppt.save_presentation(pres, str(path))
        assert path.exists()


class TestSetThemeFont:
    """Tests for set_theme_font function."""

    def test_set_heading_font(self, tmp_path):
        """Test setting the heading font."""
        pres = ppt.create_presentation()
        ppt.add_slide(pres, layout="Title Slide")  # Ensure theme is loaded

        ppt.set_theme_font(pres, "major", "Arial")

        path = tmp_path / "custom_font.pptx"
        ppt.save_presentation(pres, str(path))
        assert path.exists()

    def test_set_body_font(self, tmp_path):
        """Test setting the body font."""
        pres = ppt.create_presentation()
        ppt.add_slide(pres, layout="Title Slide")  # Ensure theme is loaded

        ppt.set_theme_font(pres, "minor", "Calibri")

        path = tmp_path / "body_font.pptx"
        ppt.save_presentation(pres, str(path))
        assert path.exists()


class TestApplyThemeColors:
    """Tests for apply_theme_colors function."""

    def test_apply_brand_colors(self, tmp_path):
        """Test applying a set of brand colors."""
        pres = ppt.create_presentation()
        ppt.add_slide(pres, layout="Title Slide")  # Ensure theme is loaded

        ppt.apply_theme_colors(
            pres,
            {
                "accent1": "#003366",
                "accent2": "#0066CC",
                "accent3": "#3399FF",
                "accent4": "#66CCFF",
            },
        )

        path = tmp_path / "brand_colors.pptx"
        ppt.save_presentation(pres, str(path))
        assert path.exists()
