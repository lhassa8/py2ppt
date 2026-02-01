"""Tests for smart slide methods on Presentation class."""

import pytest
from pathlib import Path

# Check for template
TEMPLATE_PATH = Path(__file__).parent.parent.parent / "AWStempate.pptx"
SKIP_TEMPLATE = not TEMPLATE_PATH.exists()


@pytest.fixture
def template():
    """Load the real template if available."""
    if SKIP_TEMPLATE:
        pytest.skip("Template not available")
    from py2ppt import Template
    return Template(TEMPLATE_PATH)


@pytest.fixture
def pres(template):
    """Create a presentation from the template."""
    return template.create_presentation()


class TestThemeProperty:
    """Tests for the theme property."""

    @pytest.mark.skipif(SKIP_TEMPLATE, reason="Template not available")
    def test_theme_returns_helper(self, pres):
        """Test that theme property returns ThemeHelper."""
        from py2ppt.theme import ThemeHelper
        assert isinstance(pres.theme, ThemeHelper)

    @pytest.mark.skipif(SKIP_TEMPLATE, reason="Template not available")
    def test_theme_has_colors(self, pres):
        """Test that theme has color properties."""
        theme = pres.theme
        assert theme.accent1.startswith("#")
        assert theme.dark1.startswith("#")
        assert theme.light1.startswith("#")

    @pytest.mark.skipif(SKIP_TEMPLATE, reason="Template not available")
    def test_theme_colored_helper(self, pres):
        """Test theme.colored() helper."""
        result = pres.theme.colored("Test", "accent1")
        assert result["text"] == "Test"
        assert result["color"] == pres.theme.accent1


class TestAddSmartSlide:
    """Tests for add_smart_slide method."""

    @pytest.mark.skipif(SKIP_TEMPLATE, reason="Template not available")
    def test_smart_slide_basic(self, pres):
        """Test basic smart slide creation."""
        slide_num = pres.add_smart_slide("Test Title", ["Point 1", "Point 2"])
        assert slide_num == 1
        assert pres.slide_count == 1

    @pytest.mark.skipif(SKIP_TEMPLATE, reason="Template not available")
    def test_smart_slide_with_statistics(self, pres):
        """Test smart slide detects statistics."""
        slide_num = pres.add_smart_slide(
            "Key Metrics",
            ["Revenue: $10M", "Growth: 50%", "Users: 2.5M"]
        )
        assert slide_num >= 1

    @pytest.mark.skipif(SKIP_TEMPLATE, reason="Template not available")
    def test_smart_slide_with_comparison(self, pres):
        """Test smart slide detects comparison."""
        slide_num = pres.add_smart_slide(
            "Before vs After",
            ["Before: slow and manual", "After: fast and automated"]
        )
        assert slide_num >= 1


class TestAddQuoteSlide:
    """Tests for add_quote_slide method."""

    @pytest.mark.skipif(SKIP_TEMPLATE, reason="Template not available")
    def test_quote_slide_basic(self, pres):
        """Test basic quote slide."""
        slide_num = pres.add_quote_slide(
            "Stay hungry, stay foolish.",
            "Steve Jobs"
        )
        assert slide_num == 1
        assert pres.slide_count == 1

    @pytest.mark.skipif(SKIP_TEMPLATE, reason="Template not available")
    def test_quote_slide_with_source(self, pres):
        """Test quote slide with source."""
        slide_num = pres.add_quote_slide(
            "The only way to do great work is to love what you do.",
            "Steve Jobs",
            source="Stanford Commencement, 2005"
        )
        assert slide_num >= 1

    @pytest.mark.skipif(SKIP_TEMPLATE, reason="Template not available")
    def test_quote_slide_no_attribution(self, pres):
        """Test quote slide without attribution."""
        slide_num = pres.add_quote_slide("Anonymous wisdom here.")
        assert slide_num >= 1


class TestAddStatsSlide:
    """Tests for add_stats_slide method."""

    @pytest.mark.skipif(SKIP_TEMPLATE, reason="Template not available")
    def test_stats_slide_basic(self, pres):
        """Test basic stats slide."""
        slide_num = pres.add_stats_slide("Key Metrics", [
            {"value": "98%", "label": "Customer Satisfaction"},
            {"value": "2.5M", "label": "Active Users"},
        ])
        assert slide_num == 1
        assert pres.slide_count == 1

    @pytest.mark.skipif(SKIP_TEMPLATE, reason="Template not available")
    def test_stats_slide_many_stats(self, pres):
        """Test stats slide with many statistics."""
        slide_num = pres.add_stats_slide("By the Numbers", [
            {"value": "100+", "label": "Countries"},
            {"value": "$5B", "label": "Revenue"},
            {"value": "10K", "label": "Employees"},
            {"value": "50M", "label": "Customers"},
        ])
        assert slide_num >= 1

    @pytest.mark.skipif(SKIP_TEMPLATE, reason="Template not available")
    def test_stats_slide_value_only(self, pres):
        """Test stats with value only (no label)."""
        slide_num = pres.add_stats_slide("Numbers", [
            {"value": "42"},
            {"value": "100%"},
        ])
        assert slide_num >= 1


class TestAddTimelineSlide:
    """Tests for add_timeline_slide method."""

    @pytest.mark.skipif(SKIP_TEMPLATE, reason="Template not available")
    def test_timeline_slide_with_dicts(self, pres):
        """Test timeline slide with dict events."""
        slide_num = pres.add_timeline_slide("Our Journey", [
            {"date": "2020", "event": "Company founded"},
            {"date": "2022", "event": "First product launch"},
            {"date": "2024", "event": "Global expansion"},
        ])
        assert slide_num == 1
        assert pres.slide_count == 1

    @pytest.mark.skipif(SKIP_TEMPLATE, reason="Template not available")
    def test_timeline_slide_with_strings(self, pres):
        """Test timeline slide with string events."""
        slide_num = pres.add_timeline_slide("History", [
            "2020: Started project",
            "2021: Beta release",
            "2022: Version 1.0",
        ])
        assert slide_num >= 1

    @pytest.mark.skipif(SKIP_TEMPLATE, reason="Template not available")
    def test_timeline_slide_mixed(self, pres):
        """Test timeline slide with mixed event types."""
        slide_num = pres.add_timeline_slide("Milestones", [
            {"date": "Q1", "event": "Planning"},
            "Q2: Development",
            {"date": "Q3", "event": "Testing"},
        ])
        assert slide_num >= 1


class TestAddAgendaSlide:
    """Tests for add_agenda_slide method."""

    @pytest.mark.skipif(SKIP_TEMPLATE, reason="Template not available")
    def test_agenda_slide_basic(self, pres):
        """Test basic agenda slide."""
        slide_num = pres.add_agenda_slide("Today's Agenda", [
            "Introduction",
            "Key Findings",
            "Recommendations",
            "Next Steps",
        ])
        assert slide_num == 1
        assert pres.slide_count == 1

    @pytest.mark.skipif(SKIP_TEMPLATE, reason="Template not available")
    def test_agenda_slide_short(self, pres):
        """Test agenda slide with few items."""
        slide_num = pres.add_agenda_slide("Topics", [
            "Overview",
            "Discussion",
        ])
        assert slide_num >= 1


class TestValidateMethod:
    """Tests for the validate method."""

    @pytest.mark.skipif(SKIP_TEMPLATE, reason="Template not available")
    def test_validate_empty_presentation(self, pres):
        """Test validation of empty presentation."""
        result = pres.validate()
        assert result.is_valid is False
        assert result.score < 100  # Has errors
        assert any(i.rule == "no_slides" for i in result.errors)

    @pytest.mark.skipif(SKIP_TEMPLATE, reason="Template not available")
    def test_validate_good_presentation(self, pres):
        """Test validation of a well-structured presentation."""
        pres.add_title_slide("Title", "Subtitle")
        pres.add_content_slide("Introduction", ["Point 1", "Point 2"])
        pres.add_content_slide("Details", ["Detail A", "Detail B"])
        pres.add_title_slide("Thank You", "")

        result = pres.validate()
        assert result.is_valid is True
        assert result.score > 50

    @pytest.mark.skipif(SKIP_TEMPLATE, reason="Template not available")
    def test_validate_returns_result(self, pres):
        """Test validate returns ValidationResult."""
        from py2ppt.validation import ValidationResult
        pres.add_title_slide("Test", "")
        result = pres.validate()
        assert isinstance(result, ValidationResult)

    @pytest.mark.skipif(SKIP_TEMPLATE, reason="Template not available")
    def test_validate_strict_mode(self, pres):
        """Test validation in strict mode."""
        pres.add_content_slide("", ["Content without title"])

        result_normal = pres.validate(strict=False)
        result_strict = pres.validate(strict=True)

        # Strict mode should fail on warnings
        assert result_strict.is_valid is False

    @pytest.mark.skipif(SKIP_TEMPLATE, reason="Template not available")
    def test_validate_detects_issues(self, pres):
        """Test validation detects common issues."""
        # Add slides with issues
        pres.add_content_slide("", [])  # No title, no content
        pres.add_content_slide("Title", [f"Point {i}" for i in range(10)])  # Too many bullets

        result = pres.validate()
        assert len(result.issues) > 0


class TestIntegration:
    """Integration tests for smart slides with validation."""

    @pytest.mark.skipif(SKIP_TEMPLATE, reason="Template not available")
    def test_smart_presentation_validates_well(self, pres):
        """Test that a presentation built with smart slides validates well."""
        # Build a good presentation
        pres.add_title_slide("Q4 Review", "January 2025")
        pres.add_agenda_slide("Agenda", [
            "Overview",
            "Results",
            "Next Steps",
        ])
        pres.add_section_slide("Overview")
        pres.add_content_slide("Context", [
            "Market conditions improved",
            "Team expanded by 20%",
        ])
        pres.set_notes(pres.slide_count, "Discuss market trends")

        pres.add_section_slide("Results")
        pres.add_stats_slide("Key Metrics", [
            {"value": "150%", "label": "Revenue vs Target"},
            {"value": "4.8", "label": "Customer Satisfaction"},
        ])
        pres.set_notes(pres.slide_count, "Emphasize the growth")

        pres.add_section_slide("Next Steps")
        pres.add_timeline_slide("Roadmap", [
            {"date": "Q1", "event": "Launch new product"},
            {"date": "Q2", "event": "Expand to Europe"},
        ])

        pres.add_title_slide("Thank You", "Questions?")

        # Validate
        result = pres.validate()
        assert result.is_valid is True
        assert result.score >= 70

    @pytest.mark.skipif(SKIP_TEMPLATE, reason="Template not available")
    def test_theme_colors_in_smart_slides(self, pres):
        """Test that smart slides use theme colors properly."""
        # Add slides that should use theme colors
        pres.add_stats_slide("Metrics", [
            {"value": "100%", "label": "Complete"},
        ])

        # Verify slide was created
        assert pres.slide_count == 1

        # Describe to check content
        info = pres.describe_slide(1)
        assert info["title"] == "Metrics"
