"""Tests for shape text functionality."""

import py2ppt as ppt


class TestSetShapeText:
    """Tests for set_shape_text function."""

    def test_set_text_on_shape(self):
        """Test setting text on a shape."""
        pres = ppt.create_presentation()
        ppt.add_slide(pres, layout="Blank")
        shape_id = ppt.add_shape(
            pres, 1, "rectangle", left="1in", top="1in", width="2in", height="1in"
        )

        result = ppt.set_shape_text(pres, 1, shape_id, "Hello World")

        assert result is True
        text = ppt.get_shape_text(pres, 1, shape_id)
        assert text == "Hello World"

    def test_set_text_with_bold(self):
        """Test setting bold text on a shape."""
        pres = ppt.create_presentation()
        ppt.add_slide(pres, layout="Blank")
        shape_id = ppt.add_shape(
            pres, 1, "oval", left="1in", top="1in", width="2in", height="1in"
        )

        result = ppt.set_shape_text(pres, 1, shape_id, "Bold Text", bold=True)

        assert result is True

    def test_set_text_with_italic(self):
        """Test setting italic text on a shape."""
        pres = ppt.create_presentation()
        ppt.add_slide(pres, layout="Blank")
        shape_id = ppt.add_shape(
            pres, 1, "rectangle", left="1in", top="1in", width="2in", height="1in"
        )

        result = ppt.set_shape_text(pres, 1, shape_id, "Italic Text", italic=True)

        assert result is True

    def test_set_text_with_font_size(self):
        """Test setting text with custom font size."""
        pres = ppt.create_presentation()
        ppt.add_slide(pres, layout="Blank")
        shape_id = ppt.add_shape(
            pres, 1, "rectangle", left="1in", top="1in", width="2in", height="1in"
        )

        result = ppt.set_shape_text(pres, 1, shape_id, "Large", font_size=24)

        assert result is True

    def test_set_text_with_color(self):
        """Test setting text with custom color."""
        pres = ppt.create_presentation()
        ppt.add_slide(pres, layout="Blank")
        shape_id = ppt.add_shape(
            pres, 1, "rectangle", left="1in", top="1in", width="2in", height="1in"
        )

        result = ppt.set_shape_text(pres, 1, shape_id, "Red", color="#FF0000")

        assert result is True

    def test_set_text_with_alignment(self):
        """Test setting text with alignment."""
        pres = ppt.create_presentation()
        ppt.add_slide(pres, layout="Blank")
        shape_id = ppt.add_shape(
            pres, 1, "rectangle", left="1in", top="1in", width="2in", height="1in"
        )

        result = ppt.set_shape_text(pres, 1, shape_id, "Left", align="left")

        assert result is True

    def test_set_text_replaces_existing(self):
        """Test that setting text replaces existing text."""
        pres = ppt.create_presentation()
        ppt.add_slide(pres, layout="Blank")
        shape_id = ppt.add_shape(
            pres, 1, "rectangle", left="1in", top="1in", width="2in", height="1in"
        )

        ppt.set_shape_text(pres, 1, shape_id, "Original")
        ppt.set_shape_text(pres, 1, shape_id, "Replaced")

        text = ppt.get_shape_text(pres, 1, shape_id)
        assert text == "Replaced"

    def test_set_text_invalid_shape(self):
        """Test setting text on non-existent shape."""
        pres = ppt.create_presentation()
        ppt.add_slide(pres, layout="Blank")

        result = ppt.set_shape_text(pres, 1, 9999, "Test")

        assert result is False


class TestGetShapeText:
    """Tests for get_shape_text function."""

    def test_get_text_from_shape(self):
        """Test getting text from a shape."""
        pres = ppt.create_presentation()
        ppt.add_slide(pres, layout="Blank")
        shape_id = ppt.add_shape(
            pres, 1, "rectangle", left="1in", top="1in", width="2in", height="1in"
        )
        ppt.set_shape_text(pres, 1, shape_id, "Test Content")

        text = ppt.get_shape_text(pres, 1, shape_id)

        assert text == "Test Content"

    def test_get_text_empty_shape(self):
        """Test getting text from shape without text."""
        pres = ppt.create_presentation()
        ppt.add_slide(pres, layout="Blank")
        shape_id = ppt.add_shape(
            pres, 1, "rectangle", left="1in", top="1in", width="2in", height="1in"
        )

        text = ppt.get_shape_text(pres, 1, shape_id)

        # Shape without text set returns empty string or None
        assert text == "" or text is None

    def test_get_text_invalid_shape(self):
        """Test getting text from non-existent shape."""
        pres = ppt.create_presentation()
        ppt.add_slide(pres, layout="Blank")

        result = ppt.get_shape_text(pres, 1, 9999)

        assert result is None

    def test_get_text_from_text_box(self):
        """Test getting text from a text box."""
        pres = ppt.create_presentation()
        ppt.add_slide(pres, layout="Blank")
        shape_id = ppt.add_text_box(
            pres, 1, "Box Content", left="1in", top="1in", width="4in", height="1in"
        )

        text = ppt.get_shape_text(pres, 1, shape_id)

        assert text == "Box Content"


class TestAppendShapeText:
    """Tests for append_shape_text function."""

    def test_append_text_same_paragraph(self):
        """Test appending text to the same paragraph."""
        pres = ppt.create_presentation()
        ppt.add_slide(pres, layout="Blank")
        shape_id = ppt.add_shape(
            pres, 1, "rectangle", left="1in", top="1in", width="2in", height="1in"
        )
        ppt.set_shape_text(pres, 1, shape_id, "Hello")

        result = ppt.append_shape_text(pres, 1, shape_id, " World")

        assert result is True
        text = ppt.get_shape_text(pres, 1, shape_id)
        assert text == "Hello World"

    def test_append_text_new_paragraph(self):
        """Test appending text as a new paragraph."""
        pres = ppt.create_presentation()
        ppt.add_slide(pres, layout="Blank")
        shape_id = ppt.add_shape(
            pres, 1, "rectangle", left="1in", top="1in", width="2in", height="2in"
        )
        ppt.set_shape_text(pres, 1, shape_id, "Line 1")

        result = ppt.append_shape_text(pres, 1, shape_id, "Line 2", new_paragraph=True)

        assert result is True
        text = ppt.get_shape_text(pres, 1, shape_id)
        assert "Line 1" in text
        assert "Line 2" in text

    def test_append_text_with_style(self):
        """Test appending styled text."""
        pres = ppt.create_presentation()
        ppt.add_slide(pres, layout="Blank")
        shape_id = ppt.add_shape(
            pres, 1, "rectangle", left="1in", top="1in", width="2in", height="1in"
        )
        ppt.set_shape_text(pres, 1, shape_id, "Normal")

        result = ppt.append_shape_text(
            pres, 1, shape_id, " Bold", bold=True, font_size=16
        )

        assert result is True

    def test_append_text_invalid_shape(self):
        """Test appending text to non-existent shape."""
        pres = ppt.create_presentation()
        ppt.add_slide(pres, layout="Blank")

        result = ppt.append_shape_text(pres, 1, 9999, "Test")

        assert result is False

    def test_append_to_empty_shape(self):
        """Test appending text to shape without existing text."""
        pres = ppt.create_presentation()
        ppt.add_slide(pres, layout="Blank")
        shape_id = ppt.add_shape(
            pres, 1, "rectangle", left="1in", top="1in", width="2in", height="1in"
        )

        result = ppt.append_shape_text(pres, 1, shape_id, "First Text")

        assert result is True
        text = ppt.get_shape_text(pres, 1, shape_id)
        assert text == "First Text"
