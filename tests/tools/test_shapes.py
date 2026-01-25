"""Tests for shape functionality."""


import py2ppt as ppt


class TestAddShape:
    """Tests for add_shape function."""

    def test_add_rectangle(self, tmp_path):
        """Test adding a rectangle shape."""
        pres = ppt.create_presentation()
        ppt.add_slide(pres, layout="Blank")

        shape_id = ppt.add_shape(
            pres,
            1,
            "rectangle",
            left="1in",
            top="1in",
            width="3in",
            height="2in",
            fill="#0066CC",
            text="Box 1",
        )

        assert isinstance(shape_id, int)
        path = tmp_path / "rectangle.pptx"
        ppt.save_presentation(pres, str(path))
        assert path.exists()

    def test_add_ellipse(self, tmp_path):
        """Test adding an ellipse shape."""
        pres = ppt.create_presentation()
        ppt.add_slide(pres, layout="Blank")

        shape_id = ppt.add_shape(
            pres,
            1,
            "ellipse",
            left="2in",
            top="2in",
            width="2in",
            height="2in",
            fill="accent1",
        )

        assert isinstance(shape_id, int)
        path = tmp_path / "ellipse.pptx"
        ppt.save_presentation(pres, str(path))
        assert path.exists()

    def test_add_arrow(self, tmp_path):
        """Test adding an arrow shape."""
        pres = ppt.create_presentation()
        ppt.add_slide(pres, layout="Blank")

        shape_id = ppt.add_shape(
            pres,
            1,
            "arrow_right",
            left="1in",
            top="3in",
            width="4in",
            height="1in",
            fill="#FF0000",
        )

        assert isinstance(shape_id, int)
        path = tmp_path / "arrow.pptx"
        ppt.save_presentation(pres, str(path))
        assert path.exists()

    def test_add_shape_with_gradient(self, tmp_path):
        """Test adding a shape with gradient fill."""
        pres = ppt.create_presentation()
        ppt.add_slide(pres, layout="Blank")

        shape_id = ppt.add_shape(
            pres,
            1,
            "rounded_rectangle",
            left="1in",
            top="1in",
            width="4in",
            height="2in",
            fill={"type": "gradient", "colors": ["#FF0000", "#FFFF00"]},
        )

        assert isinstance(shape_id, int)
        path = tmp_path / "gradient_shape.pptx"
        ppt.save_presentation(pres, str(path))
        assert path.exists()

    def test_add_shape_no_fill(self, tmp_path):
        """Test adding a shape with no fill."""
        pres = ppt.create_presentation()
        ppt.add_slide(pres, layout="Blank")

        shape_id = ppt.add_shape(
            pres,
            1,
            "rectangle",
            left="1in",
            top="1in",
            width="3in",
            height="2in",
            fill="none",
            outline="#000000",
        )

        assert isinstance(shape_id, int)
        path = tmp_path / "no_fill_shape.pptx"
        ppt.save_presentation(pres, str(path))
        assert path.exists()


class TestAddConnector:
    """Tests for add_connector function."""

    def test_add_straight_connector(self, tmp_path):
        """Test adding a straight connector."""
        pres = ppt.create_presentation()
        ppt.add_slide(pres, layout="Blank")

        connector_id = ppt.add_connector(
            pres,
            1,
            start_x="1in",
            start_y="1in",
            end_x="5in",
            end_y="3in",
        )

        assert isinstance(connector_id, int)
        path = tmp_path / "straight_connector.pptx"
        ppt.save_presentation(pres, str(path))
        assert path.exists()

    def test_add_arrow_connector(self, tmp_path):
        """Test adding a connector with arrows."""
        pres = ppt.create_presentation()
        ppt.add_slide(pres, layout="Blank")

        connector_id = ppt.add_connector(
            pres,
            1,
            start_x="1in",
            start_y="2in",
            end_x="5in",
            end_y="2in",
            end_arrow="triangle",
            color="#FF0000",
        )

        assert isinstance(connector_id, int)
        path = tmp_path / "arrow_connector.pptx"
        ppt.save_presentation(pres, str(path))
        assert path.exists()


class TestDeleteShape:
    """Tests for delete_shape function."""

    def test_delete_shape(self, tmp_path):
        """Test deleting a shape."""
        pres = ppt.create_presentation()
        ppt.add_slide(pres, layout="Blank")

        shape_id = ppt.add_shape(
            pres,
            1,
            "rectangle",
            left="1in",
            top="1in",
            width="2in",
            height="1in",
        )

        result = ppt.delete_shape(pres, 1, shape_id)
        assert result is True

        path = tmp_path / "deleted_shape.pptx"
        ppt.save_presentation(pres, str(path))
        assert path.exists()

    def test_delete_nonexistent_shape(self):
        """Test deleting a shape that doesn't exist."""
        pres = ppt.create_presentation()
        ppt.add_slide(pres, layout="Blank")

        result = ppt.delete_shape(pres, 1, 99999)
        assert result is False


class TestMoveShape:
    """Tests for move_shape function."""

    def test_move_shape(self, tmp_path):
        """Test moving a shape."""
        pres = ppt.create_presentation()
        ppt.add_slide(pres, layout="Blank")

        shape_id = ppt.add_shape(
            pres,
            1,
            "rectangle",
            left="1in",
            top="1in",
            width="2in",
            height="1in",
        )

        result = ppt.move_shape(pres, 1, shape_id, left="3in", top="3in")
        assert result is True

        path = tmp_path / "moved_shape.pptx"
        ppt.save_presentation(pres, str(path))
        assert path.exists()


class TestResizeShape:
    """Tests for resize_shape function."""

    def test_resize_shape(self, tmp_path):
        """Test resizing a shape."""
        pres = ppt.create_presentation()
        ppt.add_slide(pres, layout="Blank")

        shape_id = ppt.add_shape(
            pres,
            1,
            "rectangle",
            left="1in",
            top="1in",
            width="2in",
            height="1in",
        )

        result = ppt.resize_shape(pres, 1, shape_id, width="4in", height="2in")
        assert result is True

        path = tmp_path / "resized_shape.pptx"
        ppt.save_presentation(pres, str(path))
        assert path.exists()
