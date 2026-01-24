"""Tests for presentation tools."""

import os

import pytest

import py2ppt as ppt


class TestCreatePresentation:
    """Tests for create_presentation."""

    def test_create_blank(self):
        pres = ppt.create_presentation()
        assert pres is not None
        assert pres.slide_count == 0

    def test_has_layouts(self):
        pres = ppt.create_presentation()
        layouts = ppt.list_layouts(pres)
        assert len(layouts) > 0

    def test_has_theme_colors(self):
        pres = ppt.create_presentation()
        colors = ppt.get_theme_colors(pres)
        assert "accent1" in colors


class TestSavePresentation:
    """Tests for save_presentation."""

    def test_save_and_reopen(self, temp_pptx_path):
        # Create and save
        pres = ppt.create_presentation()
        ppt.add_slide(pres, layout="Title Slide")
        ppt.set_title(pres, 1, "Test Title")
        ppt.save_presentation(pres, temp_pptx_path)

        assert os.path.exists(temp_pptx_path)

        # Reopen
        pres2 = ppt.open_presentation(temp_pptx_path)
        assert pres2.slide_count == 1


class TestOpenPresentation:
    """Tests for open_presentation."""

    def test_open_saved_presentation(self, temp_pptx_path):
        # Create original
        pres = ppt.create_presentation()
        ppt.add_slide(pres, layout="Title Slide")
        ppt.set_title(pres, 1, "Original Title")
        ppt.save_presentation(pres, temp_pptx_path)

        # Open and verify
        pres2 = ppt.open_presentation(temp_pptx_path)
        slide = pres2.get_slide(1)
        assert slide.get_title() == "Original Title"
