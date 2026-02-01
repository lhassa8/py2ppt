"""Tests for theme helper module."""

import pytest
from unittest.mock import MagicMock

from py2ppt.theme import ThemeHelper


@pytest.fixture
def mock_template():
    """Create a mock template with theme colors and fonts."""
    template = MagicMock()
    template.colors = {
        "accent1": "#41B3FF",
        "accent2": "#FF6B6B",
        "accent3": "#4ECDC4",
        "accent4": "#45B7D1",
        "accent5": "#96CEB4",
        "accent6": "#FFEAA7",
        "dk1": "#2D3436",
        "dk2": "#636E72",
        "lt1": "#FFFFFF",
        "lt2": "#DFE6E9",
        "hlink": "#0984E3",
    }
    template.fonts = {
        "heading": "Arial Black",
        "body": "Arial",
    }
    return template


@pytest.fixture
def theme(mock_template):
    """Create a ThemeHelper instance."""
    return ThemeHelper(mock_template)


class TestColorProperties:
    """Tests for theme color properties."""

    def test_accent1(self, theme):
        """Test accent1 color."""
        assert theme.accent1 == "#41B3FF"

    def test_accent2(self, theme):
        """Test accent2 color."""
        assert theme.accent2 == "#FF6B6B"

    def test_accent3(self, theme):
        """Test accent3 color."""
        assert theme.accent3 == "#4ECDC4"

    def test_accent4(self, theme):
        """Test accent4 color."""
        assert theme.accent4 == "#45B7D1"

    def test_accent5(self, theme):
        """Test accent5 color."""
        assert theme.accent5 == "#96CEB4"

    def test_accent6(self, theme):
        """Test accent6 color."""
        assert theme.accent6 == "#FFEAA7"

    def test_dark1(self, theme):
        """Test dark1 color."""
        assert theme.dark1 == "#2D3436"

    def test_dark2(self, theme):
        """Test dark2 color."""
        assert theme.dark2 == "#636E72"

    def test_light1(self, theme):
        """Test light1 color."""
        assert theme.light1 == "#FFFFFF"

    def test_light2(self, theme):
        """Test light2 color."""
        assert theme.light2 == "#DFE6E9"

    def test_hyperlink(self, theme):
        """Test hyperlink color."""
        assert theme.hyperlink == "#0984E3"

    def test_accent_by_number(self, theme):
        """Test accent(n) method."""
        assert theme.accent(1) == "#41B3FF"
        assert theme.accent(2) == "#FF6B6B"
        assert theme.accent(6) == "#FFEAA7"
        # Test wrap-around
        assert theme.accent(7) == theme.accent(1)

    def test_all_colors(self, theme):
        """Test all_colors property."""
        colors = theme.all_colors
        assert "accent1" in colors
        assert "dk1" in colors
        assert "lt1" in colors


class TestFontProperties:
    """Tests for theme font properties."""

    def test_heading_font(self, theme):
        """Test heading font."""
        assert theme.heading_font == "Arial Black"

    def test_body_font(self, theme):
        """Test body font."""
        assert theme.body_font == "Arial"

    def test_all_fonts(self, theme):
        """Test all_fonts property."""
        fonts = theme.all_fonts
        assert fonts["heading"] == "Arial Black"
        assert fonts["body"] == "Arial"


class TestFormattingHelpers:
    """Tests for theme formatting helper methods."""

    def test_colored_with_theme_color(self, theme):
        """Test colored() with theme color name."""
        result = theme.colored("Important", "accent1")
        assert result["text"] == "Important"
        assert result["color"] == "#41B3FF"

    def test_colored_with_hex_color(self, theme):
        """Test colored() with hex color string."""
        result = theme.colored("Custom", "#FF0000")
        assert result["text"] == "Custom"
        assert result["color"] == "#FF0000"

    def test_colored_default_accent1(self, theme):
        """Test colored() defaults to accent1."""
        result = theme.colored("Default")
        assert result["color"] == "#41B3FF"

    def test_colored_with_extra_kwargs(self, theme):
        """Test colored() with additional kwargs."""
        result = theme.colored("Bold Red", "accent2", bold=True)
        assert result["text"] == "Bold Red"
        assert result["color"] == "#FF6B6B"
        assert result["bold"] is True

    def test_bold(self, theme):
        """Test bold() method."""
        result = theme.bold("Key point")
        assert result["text"] == "Key point"
        assert result["bold"] is True

    def test_bold_with_extra_kwargs(self, theme):
        """Test bold() with additional kwargs."""
        result = theme.bold("Big bold", font_size=24)
        assert result["bold"] is True
        assert result["font_size"] == 24

    def test_italic(self, theme):
        """Test italic() method."""
        result = theme.italic("Emphasized")
        assert result["text"] == "Emphasized"
        assert result["italic"] is True

    def test_underline(self, theme):
        """Test underline() method."""
        result = theme.underline("Underscored")
        assert result["text"] == "Underscored"
        assert result["underline"] is True

    def test_bold_colored(self, theme):
        """Test bold_colored() method."""
        result = theme.bold_colored("Highlight", "accent2")
        assert result["text"] == "Highlight"
        assert result["bold"] is True
        assert result["color"] == "#FF6B6B"

    def test_link(self, theme):
        """Test link() method."""
        result = theme.link("Click here", "https://example.com")
        assert result["text"] == "Click here"
        assert result["hyperlink"] == "https://example.com"
        assert result["color"] == "#0984E3"  # hyperlink color

    def test_heading(self, theme):
        """Test heading() method."""
        result = theme.heading("Section Title")
        assert result["text"] == "Section Title"
        assert result["bold"] is True
        assert result["font_family"] == "Arial Black"

    def test_sized(self, theme):
        """Test sized() method."""
        result = theme.sized("Big text", 36)
        assert result["text"] == "Big text"
        assert result["font_size"] == 36

    def test_sized_with_extra_kwargs(self, theme):
        """Test sized() with additional kwargs."""
        result = theme.sized("Big bold", 36, bold=True)
        assert result["font_size"] == 36
        assert result["bold"] is True


class TestCompositeFormatters:
    """Tests for composite formatting methods."""

    def test_label_value(self, theme):
        """Test label_value() method."""
        result = theme.label_value("Status:", "Complete")
        assert len(result) == 2
        assert result[0]["text"] == "Status: "
        assert result[0]["bold"] is True
        assert result[1]["text"] == "Complete"

    def test_label_value_with_color(self, theme):
        """Test label_value() with value color."""
        result = theme.label_value("Status:", "Complete", value_color="accent1")
        assert result[1]["color"] == "#41B3FF"

    def test_label_value_no_bold(self, theme):
        """Test label_value() without bold label."""
        result = theme.label_value("Note:", "Info", label_bold=False)
        assert "bold" not in result[0] or result[0].get("bold") is not True

    def test_numbered(self, theme):
        """Test numbered() method."""
        result = theme.numbered(1, "First step")
        assert len(result) == 2
        assert result[0]["text"] == "1. "
        assert result[0]["bold"] is True
        assert result[0]["color"] == "#41B3FF"  # accent1
        assert result[1]["text"] == "First step"

    def test_numbered_with_custom_color(self, theme):
        """Test numbered() with custom color."""
        result = theme.numbered(2, "Second step", number_color="accent2")
        assert result[0]["color"] == "#FF6B6B"

    def test_numbered_with_string_number(self, theme):
        """Test numbered() with string number."""
        result = theme.numbered("A", "Option A")
        assert result[0]["text"] == "A. "


class TestRepr:
    """Tests for __repr__ method."""

    def test_repr(self, theme):
        """Test string representation."""
        r = repr(theme)
        assert "ThemeHelper" in r
        assert "accent1" in r
        assert "heading" in r


class TestDefaultValues:
    """Tests for default fallback values."""

    def test_default_colors_when_missing(self):
        """Test default colors when template colors are empty."""
        template = MagicMock()
        template.colors = {}
        template.fonts = {}
        theme = ThemeHelper(template)

        # Should return defaults
        assert theme.accent1 == "#4472C4"  # PowerPoint default blue
        assert theme.dark1 == "#000000"
        assert theme.light1 == "#FFFFFF"

    def test_default_fonts_when_missing(self):
        """Test default fonts when template fonts are empty."""
        template = MagicMock()
        template.colors = {}
        template.fonts = {}
        theme = ThemeHelper(template)

        assert theme.heading_font == "Calibri Light"
        assert theme.body_font == "Calibri"
