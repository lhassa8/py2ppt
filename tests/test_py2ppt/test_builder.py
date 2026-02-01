"""Tests for presentation builder module."""

import pytest
from unittest.mock import MagicMock, patch

from py2ppt.builder import (
    SlideSpec,
    SectionSpec,
    PresentationSpec,
    build_presentation,
    build_from_outline,
    _dict_to_slide_spec,
    _dict_to_section_spec,
    _dict_to_presentation_spec,
)


class TestSlideSpec:
    """Tests for SlideSpec dataclass."""

    def test_create_basic(self):
        """Test creating a basic slide spec."""
        spec = SlideSpec(title="Test Title")
        assert spec.title == "Test Title"
        assert spec.content is None
        assert spec.slide_type is None
        assert spec.layout is None
        assert spec.notes == ""
        assert spec.extra == {}

    def test_create_with_all_fields(self):
        """Test creating a slide spec with all fields."""
        spec = SlideSpec(
            title="Full Spec",
            content=["Point 1", "Point 2"],
            slide_type="comparison",
            layout="two_column",
            notes="Speaker notes here",
            extra={"left_heading": "A", "right_heading": "B"},
        )
        assert spec.title == "Full Spec"
        assert len(spec.content) == 2
        assert spec.slide_type == "comparison"
        assert spec.extra["left_heading"] == "A"


class TestSectionSpec:
    """Tests for SectionSpec dataclass."""

    def test_create_basic(self):
        """Test creating a basic section spec."""
        spec = SectionSpec(title="Introduction")
        assert spec.title == "Introduction"
        assert spec.slides == []
        assert spec.include_divider is True

    def test_create_with_slides(self):
        """Test creating a section with slides."""
        slides = [
            SlideSpec(title="Slide 1"),
            SlideSpec(title="Slide 2"),
        ]
        spec = SectionSpec(title="Section 1", slides=slides)
        assert len(spec.slides) == 2

    def test_no_divider(self):
        """Test section without divider."""
        spec = SectionSpec(title="Hidden", include_divider=False)
        assert spec.include_divider is False


class TestPresentationSpec:
    """Tests for PresentationSpec dataclass."""

    def test_create_basic(self):
        """Test creating a basic presentation spec."""
        spec = PresentationSpec(title="My Presentation")
        assert spec.title == "My Presentation"
        assert spec.subtitle == ""
        assert spec.sections == []
        assert spec.closing_title == ""

    def test_create_full(self):
        """Test creating a full presentation spec."""
        spec = PresentationSpec(
            title="Q4 Review",
            subtitle="January 2025",
            sections=[
                SectionSpec(
                    title="Results",
                    slides=[SlideSpec(title="Revenue")],
                )
            ],
            closing_title="Thank You",
            closing_content=["Contact: email@example.com"],
        )
        assert spec.title == "Q4 Review"
        assert spec.subtitle == "January 2025"
        assert len(spec.sections) == 1
        assert spec.closing_title == "Thank You"


class TestDictConversion:
    """Tests for dict to spec conversion functions."""

    def test_dict_to_slide_spec(self):
        """Test converting dict to SlideSpec."""
        d = {
            "title": "Test Slide",
            "content": ["A", "B"],
            "slide_type": "content",
            "notes": "Notes here",
        }
        spec = _dict_to_slide_spec(d)
        assert spec.title == "Test Slide"
        assert spec.content == ["A", "B"]
        assert spec.slide_type == "content"
        assert spec.notes == "Notes here"

    def test_dict_to_slide_spec_type_alias(self):
        """Test 'type' alias for slide_type."""
        d = {"title": "Test", "type": "comparison"}
        spec = _dict_to_slide_spec(d)
        assert spec.slide_type == "comparison"

    def test_dict_to_slide_spec_extra_fields(self):
        """Test extra fields go to extra dict."""
        d = {
            "title": "Test",
            "left_heading": "A",
            "right_heading": "B",
            "custom_field": "value",
        }
        spec = _dict_to_slide_spec(d)
        assert spec.extra["left_heading"] == "A"
        assert spec.extra["right_heading"] == "B"
        assert spec.extra["custom_field"] == "value"

    def test_dict_to_section_spec(self):
        """Test converting dict to SectionSpec."""
        d = {
            "title": "Section 1",
            "slides": [
                {"title": "Slide 1"},
                {"title": "Slide 2"},
            ],
            "include_divider": False,
        }
        spec = _dict_to_section_spec(d)
        assert spec.title == "Section 1"
        assert len(spec.slides) == 2
        assert isinstance(spec.slides[0], SlideSpec)
        assert spec.include_divider is False

    def test_dict_to_presentation_spec(self):
        """Test converting dict to PresentationSpec."""
        d = {
            "title": "Presentation",
            "subtitle": "Subtitle",
            "sections": [
                {
                    "title": "Intro",
                    "slides": [{"title": "Welcome"}],
                }
            ],
            "closing_title": "Thanks",
        }
        spec = _dict_to_presentation_spec(d)
        assert spec.title == "Presentation"
        assert spec.subtitle == "Subtitle"
        assert len(spec.sections) == 1
        assert spec.closing_title == "Thanks"

    def test_closing_alias(self):
        """Test 'closing' alias for closing_title."""
        d = {"title": "Test", "closing": "Goodbye"}
        spec = _dict_to_presentation_spec(d)
        assert spec.closing_title == "Goodbye"


class TestBuildPresentation:
    """Tests for build_presentation function."""

    @pytest.fixture
    def mock_template(self):
        """Create a mock template."""
        template = MagicMock()
        pres = MagicMock()
        pres.add_title_slide = MagicMock(return_value=1)
        pres.add_section_slide = MagicMock(return_value=2)
        pres.add_content_slide = MagicMock(return_value=3)
        pres.add_comparison_slide = MagicMock(return_value=4)
        pres.add_table_slide = MagicMock(return_value=5)
        pres.add_chart_slide = MagicMock(return_value=6)
        pres.add_quote_slide = MagicMock(return_value=7)
        pres.add_stats_slide = MagicMock(return_value=8)
        pres.add_timeline_slide = MagicMock(return_value=9)
        pres.add_agenda_slide = MagicMock(return_value=10)
        pres.add_two_column_slide = MagicMock(return_value=11)
        pres.add_image_slide = MagicMock(return_value=12)
        pres.add_blank_slide = MagicMock(return_value=13)
        pres.set_notes = MagicMock()
        template.create_presentation = MagicMock(return_value=pres)
        return template

    def test_build_from_spec(self, mock_template):
        """Test building from PresentationSpec."""
        spec = PresentationSpec(
            title="Test",
            subtitle="Subtitle",
            closing_title="Thanks",
        )
        result = build_presentation(mock_template, spec)

        mock_template.create_presentation.assert_called_once()
        result.add_title_slide.assert_any_call("Test", "Subtitle", layout=None)
        result.add_title_slide.assert_any_call("Thanks", "", layout=None)

    def test_build_from_dict(self, mock_template):
        """Test building from dict."""
        spec_dict = {
            "title": "Dict Presentation",
            "subtitle": "From Dict",
        }
        result = build_presentation(mock_template, spec_dict)

        result.add_title_slide.assert_called()

    def test_build_with_sections(self, mock_template):
        """Test building with sections."""
        spec = PresentationSpec(
            title="Sectioned",
            sections=[
                SectionSpec(
                    title="Section 1",
                    slides=[SlideSpec(title="Slide 1", content=["A", "B"])],
                ),
            ],
        )
        result = build_presentation(mock_template, spec)

        result.add_section_slide.assert_called_with("Section 1", layout=None)
        result.add_content_slide.assert_called()

    def test_build_with_section_no_divider(self, mock_template):
        """Test section with include_divider=False."""
        spec = PresentationSpec(
            title="No Divider",
            sections=[
                SectionSpec(
                    title="Hidden Section",
                    slides=[SlideSpec(title="Slide")],
                    include_divider=False,
                ),
            ],
        )
        result = build_presentation(mock_template, spec)

        # Section slide should not be called for this section
        result.add_section_slide.assert_not_called()

    def test_build_with_table_slide(self, mock_template):
        """Test building with table slide type."""
        spec = PresentationSpec(
            title="Tables",
            sections=[
                SectionSpec(
                    title="Data",
                    slides=[
                        SlideSpec(
                            title="Table Slide",
                            slide_type="table",
                            extra={"headers": ["A", "B"], "rows": [[1, 2]]},
                        )
                    ],
                ),
            ],
        )
        result = build_presentation(mock_template, spec)

        result.add_table_slide.assert_called()

    def test_build_with_chart_slide(self, mock_template):
        """Test building with chart slide type."""
        spec = PresentationSpec(
            title="Charts",
            sections=[
                SectionSpec(
                    title="Data",
                    slides=[
                        SlideSpec(
                            title="Chart Slide",
                            slide_type="chart",
                            extra={
                                "chart_type": "bar",
                                "data": {"categories": ["A"], "series": [{"name": "S", "values": [1]}]},
                            },
                        )
                    ],
                ),
            ],
        )
        result = build_presentation(mock_template, spec)

        result.add_chart_slide.assert_called()

    def test_build_with_quote_slide(self, mock_template):
        """Test building with quote slide type."""
        spec = PresentationSpec(
            title="Quotes",
            sections=[
                SectionSpec(
                    title="Inspiration",
                    slides=[
                        SlideSpec(
                            title="Quote",
                            content=["Great quote here"],
                            slide_type="quote",
                            extra={"attribution": "Author"},
                        )
                    ],
                ),
            ],
        )
        result = build_presentation(mock_template, spec)

        result.add_quote_slide.assert_called()

    def test_build_with_notes(self, mock_template):
        """Test building slides with notes."""
        spec = PresentationSpec(
            title="Notes",
            sections=[
                SectionSpec(
                    title="Content",
                    slides=[
                        SlideSpec(
                            title="With Notes",
                            content=["Content"],
                            notes="Speaker notes here",
                        )
                    ],
                ),
            ],
        )
        result = build_presentation(mock_template, spec)

        result.set_notes.assert_called()

    def test_build_with_closing_content(self, mock_template):
        """Test closing slide with content."""
        spec = PresentationSpec(
            title="Test",
            closing_title="Contact",
            closing_content=["email@example.com", "twitter.com/handle"],
        )
        result = build_presentation(mock_template, spec)

        # Should use add_content_slide for closing with content
        result.add_content_slide.assert_called()


class TestBuildFromOutline:
    """Tests for build_from_outline function."""

    @pytest.fixture
    def mock_template(self):
        """Create a mock template."""
        template = MagicMock()
        pres = MagicMock()
        pres.add_title_slide = MagicMock(return_value=1)
        pres.add_section_slide = MagicMock(return_value=2)
        pres.add_content_slide = MagicMock(return_value=3)
        template.create_presentation = MagicMock(return_value=pres)
        return template

    def test_build_simple_outline(self, mock_template):
        """Test building from simple string outline."""
        outline = [
            "Introduction",
            "Main Content",
            "Conclusion",
        ]
        result = build_from_outline(mock_template, "My Talk", outline)

        mock_template.create_presentation.assert_called_once()
        result.add_title_slide.assert_called()
        assert result.add_content_slide.call_count >= 3

    def test_build_with_dict_slides(self, mock_template):
        """Test building with dict slide specs."""
        outline = [
            {"title": "Slide 1", "content": ["A", "B"]},
            {"title": "Slide 2", "content": ["C", "D"]},
        ]
        result = build_from_outline(mock_template, "Talk", outline)

        result.add_content_slide.assert_called()

    def test_build_with_sections(self, mock_template):
        """Test building with section markers."""
        outline = [
            {"title": "Part 1", "section": True},
            {"title": "Slide 1", "content": ["Content"]},
            {"title": "Part 2", "section": True},
            {"title": "Slide 2", "content": ["More content"]},
        ]
        result = build_from_outline(mock_template, "Talk", outline)

        assert result.add_section_slide.call_count == 2

    def test_build_no_auto_sections(self, mock_template):
        """Test building without auto sections."""
        outline = [
            {"title": "Part 1", "section": True},
            {"title": "Slide 1"},
        ]
        result = build_from_outline(
            mock_template, "Talk", outline, auto_sections=False
        )

        # Section markers should be treated as regular slides
        result.add_section_slide.assert_not_called()

    def test_build_with_subtitle(self, mock_template):
        """Test building with subtitle."""
        outline = ["Slide 1"]
        result = build_from_outline(
            mock_template, "Talk", outline, subtitle="Subtitle Here"
        )

        result.add_title_slide.assert_any_call("Talk", "Subtitle Here")

    def test_build_with_custom_closing(self, mock_template):
        """Test building with custom closing."""
        outline = ["Slide 1"]
        result = build_from_outline(
            mock_template, "Talk", outline, closing="Questions?"
        )

        result.add_title_slide.assert_any_call("Questions?", "")

    def test_build_no_closing(self, mock_template):
        """Test building without closing slide."""
        outline = ["Slide 1"]
        result = build_from_outline(
            mock_template, "Talk", outline, closing=""
        )

        # Should only call add_title_slide for the opening
        assert result.add_title_slide.call_count == 1


class TestAutoDetection:
    """Tests for auto-detection of slide types."""

    @pytest.fixture
    def mock_template(self):
        """Create a mock template with full slide methods."""
        template = MagicMock()
        pres = MagicMock()
        pres.add_title_slide = MagicMock(return_value=1)
        pres.add_section_slide = MagicMock(return_value=2)
        pres.add_content_slide = MagicMock(return_value=3)
        pres.add_comparison_slide = MagicMock(return_value=4)
        pres.set_notes = MagicMock()
        template.create_presentation = MagicMock(return_value=pres)
        return template

    def test_auto_detect_comparison(self, mock_template):
        """Test auto-detection of comparison content."""
        spec = PresentationSpec(
            title="Test",
            sections=[
                SectionSpec(
                    title="Compare",
                    slides=[
                        SlideSpec(
                            title="Before vs After",
                            content=["Before: old", "After: new"],
                        )
                    ],
                ),
            ],
        )

        with patch("py2ppt.builder.analyze_content") as mock_analyze:
            mock_analyze.return_value = MagicMock(
                confidence=0.8,
                recommended_slide_type="comparison",
            )
            result = build_presentation(mock_template, spec)
            # Should use comparison slide method
            result.add_comparison_slide.assert_called()
