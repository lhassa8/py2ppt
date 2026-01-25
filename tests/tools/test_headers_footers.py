"""Tests for header and footer functionality."""

import py2ppt as ppt


class TestSetFooter:
    """Tests for set_footer function."""

    def test_set_footer(self):
        """Test setting footer text."""
        pres = ppt.create_presentation()
        ppt.add_slide(pres, layout="Title and Content")

        # Should not raise
        ppt.set_footer(pres, "Confidential")

    def test_set_footer_single_slide(self):
        """Test setting footer on a specific slide."""
        pres = ppt.create_presentation()
        ppt.add_slide(pres, layout="Title and Content")
        ppt.add_slide(pres, layout="Title and Content")

        # Should not raise
        ppt.set_footer(pres, "Draft", apply_to_all=False, slide_number=1)


class TestSetSlideNumberVisibility:
    """Tests for set_slide_number_visibility function."""

    def test_show_slide_numbers(self):
        """Test enabling slide numbers."""
        pres = ppt.create_presentation()
        ppt.add_slide(pres, layout="Title and Content")

        ppt.set_slide_number_visibility(pres, visible=True)
        settings = ppt.get_header_footer_settings(pres)

        assert settings["slide_number_visible"] is True

    def test_hide_slide_numbers(self):
        """Test disabling slide numbers."""
        pres = ppt.create_presentation()
        ppt.add_slide(pres, layout="Title and Content")

        ppt.set_slide_number_visibility(pres, visible=False)
        settings = ppt.get_header_footer_settings(pres)

        assert settings["slide_number_visible"] is False

    def test_slide_numbers_start_from(self):
        """Test setting start number for slides."""
        pres = ppt.create_presentation()
        ppt.add_slide(pres, layout="Title and Content")

        ppt.set_slide_number_visibility(pres, visible=True, start_from=0)
        settings = ppt.get_header_footer_settings(pres)

        assert settings["first_slide_number"] == 0


class TestSetDateVisibility:
    """Tests for set_date_visibility function."""

    def test_show_date(self):
        """Test enabling date display."""
        pres = ppt.create_presentation()
        ppt.add_slide(pres, layout="Title and Content")

        ppt.set_date_visibility(pres, visible=True)
        settings = ppt.get_header_footer_settings(pres)

        assert settings["date_visible"] is True

    def test_hide_date(self):
        """Test disabling date display."""
        pres = ppt.create_presentation()
        ppt.add_slide(pres, layout="Title and Content")

        ppt.set_date_visibility(pres, visible=False)
        settings = ppt.get_header_footer_settings(pres)

        assert settings["date_visible"] is False


class TestGetHeaderFooterSettings:
    """Tests for get_header_footer_settings function."""

    def test_get_settings_default(self):
        """Test getting default settings."""
        pres = ppt.create_presentation()

        settings = ppt.get_header_footer_settings(pres)

        assert "footer_visible" in settings
        assert "slide_number_visible" in settings
        assert "date_visible" in settings
        assert "first_slide_number" in settings
