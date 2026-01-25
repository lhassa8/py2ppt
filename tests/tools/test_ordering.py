"""Tests for object ordering functionality."""

import py2ppt as ppt


class TestGetShapeOrder:
    """Tests for get_shape_order function."""

    def test_get_shape_order_with_shapes(self):
        """Test getting shape order from slide with shapes."""
        pres = ppt.create_presentation()
        ppt.add_slide(pres, layout="Blank")

        # Add some shapes
        ppt.add_shape(pres, 1, "rectangle", left="1in", top="1in", width="2in", height="1in")
        ppt.add_shape(pres, 1, "ellipse", left="2in", top="2in", width="2in", height="1in")

        order = ppt.get_shape_order(pres, 1)

        assert isinstance(order, list)
        assert len(order) >= 2  # At least our two shapes


class TestBringToFront:
    """Tests for bring_to_front function."""

    def test_bring_to_front(self):
        """Test bringing a shape to front."""
        pres = ppt.create_presentation()
        ppt.add_slide(pres, layout="Blank")

        shape1 = ppt.add_shape(pres, 1, "rectangle", left="1in", top="1in", width="2in", height="1in")
        ppt.add_shape(pres, 1, "ellipse", left="2in", top="2in", width="2in", height="1in")

        # Bring first shape to front
        result = ppt.bring_to_front(pres, 1, shape1)

        assert result is True


class TestSendToBack:
    """Tests for send_to_back function."""

    def test_send_to_back(self):
        """Test sending a shape to back."""
        pres = ppt.create_presentation()
        ppt.add_slide(pres, layout="Blank")

        ppt.add_shape(pres, 1, "rectangle", left="1in", top="1in", width="2in", height="1in")
        shape2 = ppt.add_shape(pres, 1, "ellipse", left="2in", top="2in", width="2in", height="1in")

        # Send second shape to back
        result = ppt.send_to_back(pres, 1, shape2)

        assert result is True


class TestBringForward:
    """Tests for bring_forward function."""

    def test_bring_forward(self):
        """Test bringing a shape forward by one level."""
        pres = ppt.create_presentation()
        ppt.add_slide(pres, layout="Blank")

        shape1 = ppt.add_shape(pres, 1, "rectangle", left="1in", top="1in", width="2in", height="1in")
        ppt.add_shape(pres, 1, "ellipse", left="2in", top="2in", width="2in", height="1in")

        result = ppt.bring_forward(pres, 1, shape1)

        assert result is True


class TestSendBackward:
    """Tests for send_backward function."""

    def test_send_backward(self):
        """Test sending a shape backward by one level."""
        pres = ppt.create_presentation()
        ppt.add_slide(pres, layout="Blank")

        ppt.add_shape(pres, 1, "rectangle", left="1in", top="1in", width="2in", height="1in")
        shape2 = ppt.add_shape(pres, 1, "ellipse", left="2in", top="2in", width="2in", height="1in")

        result = ppt.send_backward(pres, 1, shape2)

        assert result is True


class TestSetShapeOrder:
    """Tests for set_shape_order function."""

    def test_set_shape_order(self):
        """Test setting explicit shape order."""
        pres = ppt.create_presentation()
        ppt.add_slide(pres, layout="Blank")

        ppt.add_shape(pres, 1, "rectangle", left="1in", top="1in", width="2in", height="1in")
        ppt.add_shape(pres, 1, "ellipse", left="2in", top="2in", width="2in", height="1in")

        # Get current order
        order = ppt.get_shape_order(pres, 1)

        # Reverse and set
        if len(order) >= 2:
            result = ppt.set_shape_order(pres, 1, list(reversed(order)))
            assert result is True
