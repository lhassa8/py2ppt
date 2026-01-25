"""Tests for text columns functionality."""

import py2ppt as ppt


class TestSetTextColumns:
    """Tests for set_text_columns function."""

    def test_set_two_columns(self):
        """Test setting two columns on a text box."""
        pres = ppt.create_presentation()
        ppt.add_slide(pres, layout="Blank")
        shape_id = ppt.add_text_box(
            pres, 1, "Column text", left="1in", top="1in", width="4in", height="2in"
        )

        result = ppt.set_text_columns(pres, 1, shape_id, num_columns=2)

        assert result is True

    def test_set_three_columns(self):
        """Test setting three columns."""
        pres = ppt.create_presentation()
        ppt.add_slide(pres, layout="Blank")
        shape_id = ppt.add_text_box(
            pres, 1, "Text", left="1in", top="1in", width="6in", height="2in"
        )

        result = ppt.set_text_columns(pres, 1, shape_id, num_columns=3)

        assert result is True
        col_info = ppt.get_text_columns(pres, 1, shape_id)
        assert col_info["columns"] == 3

    def test_set_columns_with_spacing(self):
        """Test setting columns with custom spacing."""
        pres = ppt.create_presentation()
        ppt.add_slide(pres, layout="Blank")
        shape_id = ppt.add_text_box(
            pres, 1, "Text", left="1in", top="1in", width="4in", height="2in"
        )

        result = ppt.set_text_columns(pres, 1, shape_id, num_columns=2, spacing="1in")

        assert result is True

    def test_set_single_column(self):
        """Test resetting to single column."""
        pres = ppt.create_presentation()
        ppt.add_slide(pres, layout="Blank")
        shape_id = ppt.add_text_box(
            pres, 1, "Text", left="1in", top="1in", width="4in", height="2in"
        )

        ppt.set_text_columns(pres, 1, shape_id, num_columns=2)
        result = ppt.set_text_columns(pres, 1, shape_id, num_columns=1)

        assert result is True
        col_info = ppt.get_text_columns(pres, 1, shape_id)
        assert col_info["columns"] == 1

    def test_set_columns_invalid_shape(self):
        """Test setting columns on non-existent shape."""
        pres = ppt.create_presentation()
        ppt.add_slide(pres, layout="Blank")

        result = ppt.set_text_columns(pres, 1, 9999, num_columns=2)

        assert result is False


class TestGetTextColumns:
    """Tests for get_text_columns function."""

    def test_get_columns_default(self):
        """Test getting columns from shape with default (1 column)."""
        pres = ppt.create_presentation()
        ppt.add_slide(pres, layout="Blank")
        shape_id = ppt.add_text_box(
            pres, 1, "Text", left="1in", top="1in", width="4in", height="2in"
        )

        col_info = ppt.get_text_columns(pres, 1, shape_id)

        assert col_info is not None
        assert "columns" in col_info

    def test_get_columns_after_set(self):
        """Test getting columns after setting them."""
        pres = ppt.create_presentation()
        ppt.add_slide(pres, layout="Blank")
        shape_id = ppt.add_text_box(
            pres, 1, "Text", left="1in", top="1in", width="6in", height="2in"
        )

        ppt.set_text_columns(pres, 1, shape_id, num_columns=4)
        col_info = ppt.get_text_columns(pres, 1, shape_id)

        assert col_info["columns"] == 4

    def test_get_columns_invalid_shape(self):
        """Test getting columns from non-existent shape."""
        pres = ppt.create_presentation()
        ppt.add_slide(pres, layout="Blank")

        result = ppt.get_text_columns(pres, 1, 9999)

        assert result is None
