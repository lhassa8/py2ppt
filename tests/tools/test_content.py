"""Tests for content tools."""

import pytest

import py2ppt as ppt
from py2ppt.utils.errors import PlaceholderNotFoundError


class TestSetTitle:
    """Tests for set_title."""

    def test_set_title(self, blank_presentation):
        ppt.add_slide(blank_presentation, layout="Title Slide")
        ppt.set_title(blank_presentation, 1, "My Title")

        slide = blank_presentation.get_slide(1)
        assert slide.get_title() == "My Title"

    def test_set_title_with_styling(self, blank_presentation):
        ppt.add_slide(blank_presentation, layout="Title Slide")
        ppt.set_title(
            blank_presentation, 1, "Styled Title",
            font_size=48, bold=True, color="#FF0000"
        )

        slide = blank_presentation.get_slide(1)
        assert slide.get_title() == "Styled Title"


class TestSetSubtitle:
    """Tests for set_subtitle."""

    def test_set_subtitle(self, blank_presentation):
        ppt.add_slide(blank_presentation, layout="Title Slide")
        ppt.set_subtitle(blank_presentation, 1, "My Subtitle")

        slide = blank_presentation.get_slide(1)
        assert slide.get_subtitle() == "My Subtitle"


class TestSetBody:
    """Tests for set_body."""

    def test_set_body_with_list(self, blank_presentation):
        ppt.add_slide(blank_presentation, layout="Title and Content")
        ppt.set_body(blank_presentation, 1, ["Point 1", "Point 2", "Point 3"])

        slide = blank_presentation.get_slide(1)
        body = slide.get_body()
        assert len(body) == 3
        assert body[0] == "Point 1"

    def test_set_body_with_string(self, blank_presentation):
        ppt.add_slide(blank_presentation, layout="Title and Content")
        ppt.set_body(blank_presentation, 1, "Single point")

        slide = blank_presentation.get_slide(1)
        body = slide.get_body()
        assert len(body) == 1

    def test_set_body_with_levels(self, blank_presentation):
        ppt.add_slide(blank_presentation, layout="Title and Content")
        ppt.set_body(
            blank_presentation, 1,
            ["Main point", "Sub point", "Another main"],
            levels=[0, 1, 0]
        )

        slide = blank_presentation.get_slide(1)
        body = slide.get_body()
        assert len(body) == 3


class TestAddBullet:
    """Tests for add_bullet."""

    def test_add_bullet(self, blank_presentation):
        ppt.add_slide(blank_presentation, layout="Title and Content")
        ppt.set_body(blank_presentation, 1, ["Initial point"])
        ppt.add_bullet(blank_presentation, 1, "Added point")

        slide = blank_presentation.get_slide(1)
        body = slide.get_body()
        assert len(body) == 2
        assert body[1] == "Added point"

    def test_add_nested_bullet(self, blank_presentation):
        ppt.add_slide(blank_presentation, layout="Title and Content")
        ppt.add_bullet(blank_presentation, 1, "Main", level=0)
        ppt.add_bullet(blank_presentation, 1, "Sub", level=1)

        slide = blank_presentation.get_slide(1)
        body = slide.get_body()
        assert len(body) == 2


class TestDescribeSlide:
    """Tests for describe_slide."""

    def test_describe_slide(self, presentation_with_slides):
        info = ppt.describe_slide(presentation_with_slides, 1)

        assert "slide_number" in info
        assert info["slide_number"] == 1
        assert "placeholders" in info

    def test_describe_slide_with_content(self, presentation_with_slides):
        info = ppt.describe_slide(presentation_with_slides, 2)

        assert "placeholders" in info
        # Should have title and body
