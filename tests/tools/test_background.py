"""Tests for slide background functionality."""

import py2ppt as ppt


class TestSetSlideBackground:
    """Tests for set_slide_background function."""

    def test_set_solid_color_hex(self):
        """Test setting solid hex color background."""
        pres = ppt.create_presentation()
        ppt.add_slide(pres, layout="Title and Content")

        # Should not raise
        ppt.set_slide_background(pres, 1, color="#003366")

    def test_set_solid_color_theme(self):
        """Test setting solid theme color background."""
        pres = ppt.create_presentation()
        ppt.add_slide(pres, layout="Title and Content")

        # Should not raise
        ppt.set_slide_background(pres, 1, color="accent1")

    def test_set_solid_color_with_transparency(self):
        """Test setting solid color with transparency."""
        pres = ppt.create_presentation()
        ppt.add_slide(pres, layout="Title and Content")

        # Should not raise
        ppt.set_slide_background(pres, 1, color="#003366", transparency=50)

    def test_set_gradient_background(self):
        """Test setting gradient background."""
        pres = ppt.create_presentation()
        ppt.add_slide(pres, layout="Title and Content")

        gradient = {
            "colors": ["#000066", "#0066CC"],
            "direction": 90,
        }
        ppt.set_slide_background(pres, 1, gradient=gradient)

    def test_set_radial_gradient(self):
        """Test setting radial gradient background."""
        pres = ppt.create_presentation()
        ppt.add_slide(pres, layout="Title and Content")

        gradient = {
            "colors": ["#FFFFFF", "#000000"],
            "type": "radial",
        }
        ppt.set_slide_background(pres, 1, gradient=gradient)


class TestClearSlideBackground:
    """Tests for clear_slide_background function."""

    def test_clear_background(self):
        """Test clearing slide background."""
        pres = ppt.create_presentation()
        ppt.add_slide(pres, layout="Title and Content")

        # Set then clear
        ppt.set_slide_background(pres, 1, color="#003366")
        ppt.clear_slide_background(pres, 1)

        # Should not raise

    def test_clear_background_no_prior_background(self):
        """Test clearing background when none exists."""
        pres = ppt.create_presentation()
        ppt.add_slide(pres, layout="Title and Content")

        # Should not raise even if no background set
        ppt.clear_slide_background(pres, 1)
