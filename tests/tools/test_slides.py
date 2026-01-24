"""Tests for slide tools."""

import pytest

import py2ppt as ppt
from py2ppt.utils.errors import LayoutNotFoundError, SlideNotFoundError


class TestAddSlide:
    """Tests for add_slide."""

    def test_add_slide_by_name(self, blank_presentation):
        slide_num = ppt.add_slide(blank_presentation, layout="Title Slide")
        assert slide_num == 1
        assert blank_presentation.slide_count == 1

    def test_add_slide_by_index(self, blank_presentation):
        slide_num = ppt.add_slide(blank_presentation, layout=0)
        assert slide_num == 1

    def test_add_multiple_slides(self, blank_presentation):
        ppt.add_slide(blank_presentation, layout="Title Slide")
        ppt.add_slide(blank_presentation, layout="Title and Content")
        ppt.add_slide(blank_presentation, layout="Blank")
        assert blank_presentation.slide_count == 3

    def test_add_slide_at_position(self, blank_presentation):
        ppt.add_slide(blank_presentation, layout="Title Slide")
        ppt.add_slide(blank_presentation, layout="Blank")
        # Insert at position 2
        ppt.add_slide(blank_presentation, layout="Title and Content", position=2)
        assert blank_presentation.slide_count == 3

    def test_add_slide_fuzzy_match(self, blank_presentation):
        # Should fuzzy match "title slide"
        slide_num = ppt.add_slide(blank_presentation, layout="title")
        assert slide_num == 1

    def test_add_slide_invalid_layout(self, blank_presentation):
        with pytest.raises(LayoutNotFoundError):
            ppt.add_slide(blank_presentation, layout="NonexistentLayout12345")


class TestDeleteSlide:
    """Tests for delete_slide."""

    def test_delete_slide(self, presentation_with_slides):
        initial_count = presentation_with_slides.slide_count
        result = ppt.delete_slide(presentation_with_slides, 1)
        assert result is True
        assert presentation_with_slides.slide_count == initial_count - 1

    def test_delete_invalid_slide(self, presentation_with_slides):
        result = ppt.delete_slide(presentation_with_slides, 999)
        assert result is False


class TestGetSlideCount:
    """Tests for get_slide_count."""

    def test_empty_presentation(self, blank_presentation):
        assert ppt.get_slide_count(blank_presentation) == 0

    def test_with_slides(self, presentation_with_slides):
        assert ppt.get_slide_count(presentation_with_slides) == 2
