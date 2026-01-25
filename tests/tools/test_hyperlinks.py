"""Tests for hyperlink functionality."""

import py2ppt as ppt


class TestAddHyperlink:
    """Tests for add_hyperlink function."""

    def test_add_external_url(self):
        """Test adding external URL hyperlink."""
        pres = ppt.create_presentation()
        ppt.add_slide(pres, layout="Blank")
        shape_id = ppt.add_shape(
            pres, 1, "rectangle",
            left="1in", top="1in", width="2in", height="1in"
        )

        result = ppt.add_hyperlink(pres, 1, shape_id, url="https://example.com")

        assert result is True

    def test_add_hyperlink_with_tooltip(self):
        """Test adding hyperlink with tooltip."""
        pres = ppt.create_presentation()
        ppt.add_slide(pres, layout="Blank")
        shape_id = ppt.add_shape(
            pres, 1, "rectangle",
            left="1in", top="1in", width="2in", height="1in"
        )

        result = ppt.add_hyperlink(
            pres, 1, shape_id,
            url="https://example.com",
            tooltip="Click to visit"
        )

        assert result is True

    def test_add_internal_slide_link(self):
        """Test adding internal slide link."""
        pres = ppt.create_presentation()
        ppt.add_slide(pres, layout="Blank")
        ppt.add_slide(pres, layout="Blank")
        shape_id = ppt.add_shape(
            pres, 1, "rectangle",
            left="1in", top="1in", width="2in", height="1in"
        )

        result = ppt.add_hyperlink(pres, 1, shape_id, slide=2)

        assert result is True

    def test_add_action_next_slide(self):
        """Test adding next slide action."""
        pres = ppt.create_presentation()
        ppt.add_slide(pres, layout="Blank")
        ppt.add_slide(pres, layout="Blank")
        shape_id = ppt.add_shape(
            pres, 1, "rectangle",
            left="1in", top="1in", width="2in", height="1in"
        )

        result = ppt.add_hyperlink(pres, 1, shape_id, action="next_slide")

        assert result is True

    def test_add_action_previous_slide(self):
        """Test adding previous slide action."""
        pres = ppt.create_presentation()
        ppt.add_slide(pres, layout="Blank")
        ppt.add_slide(pres, layout="Blank")
        shape_id = ppt.add_shape(
            pres, 2, "rectangle",
            left="1in", top="1in", width="2in", height="1in"
        )

        result = ppt.add_hyperlink(pres, 2, shape_id, action="previous_slide")

        assert result is True

    def test_add_hyperlink_nonexistent_shape(self):
        """Test adding hyperlink to non-existent shape."""
        pres = ppt.create_presentation()
        ppt.add_slide(pres, layout="Blank")

        result = ppt.add_hyperlink(pres, 1, 9999, url="https://example.com")

        assert result is False

    def test_add_hyperlink_invalid_slide(self):
        """Test adding hyperlink to invalid slide."""
        pres = ppt.create_presentation()
        ppt.add_slide(pres, layout="Blank")
        shape_id = ppt.add_shape(
            pres, 1, "rectangle",
            left="1in", top="1in", width="2in", height="1in"
        )

        result = ppt.add_hyperlink(pres, 99, shape_id, url="https://example.com")

        assert result is False


class TestRemoveHyperlink:
    """Tests for remove_hyperlink function."""

    def test_remove_hyperlink(self):
        """Test removing hyperlink from shape."""
        pres = ppt.create_presentation()
        ppt.add_slide(pres, layout="Blank")
        shape_id = ppt.add_shape(
            pres, 1, "rectangle",
            left="1in", top="1in", width="2in", height="1in"
        )

        # Add then remove
        ppt.add_hyperlink(pres, 1, shape_id, url="https://example.com")
        result = ppt.remove_hyperlink(pres, 1, shape_id)

        assert result is True

    def test_remove_hyperlink_no_link(self):
        """Test removing hyperlink when none exists."""
        pres = ppt.create_presentation()
        ppt.add_slide(pres, layout="Blank")
        shape_id = ppt.add_shape(
            pres, 1, "rectangle",
            left="1in", top="1in", width="2in", height="1in"
        )

        result = ppt.remove_hyperlink(pres, 1, shape_id)

        assert result is False

    def test_remove_hyperlink_nonexistent_shape(self):
        """Test removing hyperlink from non-existent shape."""
        pres = ppt.create_presentation()
        ppt.add_slide(pres, layout="Blank")

        result = ppt.remove_hyperlink(pres, 1, 9999)

        assert result is False


class TestGetHyperlinks:
    """Tests for get_hyperlinks function."""

    def test_get_hyperlinks_empty(self):
        """Test getting hyperlinks when none exist."""
        pres = ppt.create_presentation()
        ppt.add_slide(pres, layout="Blank")

        links = ppt.get_hyperlinks(pres, 1)

        assert isinstance(links, list)
        assert len(links) == 0

    def test_get_hyperlinks_with_url(self):
        """Test getting hyperlinks with external URL."""
        pres = ppt.create_presentation()
        ppt.add_slide(pres, layout="Blank")
        shape_id = ppt.add_shape(
            pres, 1, "rectangle",
            left="1in", top="1in", width="2in", height="1in"
        )

        ppt.add_hyperlink(pres, 1, shape_id, url="https://example.com")
        links = ppt.get_hyperlinks(pres, 1)

        assert len(links) == 1
        assert links[0]["shape_id"] == shape_id
        assert links[0]["url"] == "https://example.com"

    def test_get_hyperlinks_with_tooltip(self):
        """Test getting hyperlinks with tooltip."""
        pres = ppt.create_presentation()
        ppt.add_slide(pres, layout="Blank")
        shape_id = ppt.add_shape(
            pres, 1, "rectangle",
            left="1in", top="1in", width="2in", height="1in"
        )

        ppt.add_hyperlink(
            pres, 1, shape_id,
            url="https://example.com",
            tooltip="Visit site"
        )
        links = ppt.get_hyperlinks(pres, 1)

        assert len(links) == 1
        assert links[0]["tooltip"] == "Visit site"

    def test_get_hyperlinks_multiple(self):
        """Test getting multiple hyperlinks."""
        pres = ppt.create_presentation()
        ppt.add_slide(pres, layout="Blank")

        shape1 = ppt.add_shape(
            pres, 1, "rectangle",
            left="1in", top="1in", width="2in", height="1in"
        )
        shape2 = ppt.add_shape(
            pres, 1, "ellipse",
            left="4in", top="1in", width="2in", height="1in"
        )

        ppt.add_hyperlink(pres, 1, shape1, url="https://example1.com")
        ppt.add_hyperlink(pres, 1, shape2, url="https://example2.com")

        links = ppt.get_hyperlinks(pres, 1)

        assert len(links) == 2

    def test_get_hyperlinks_invalid_slide(self):
        """Test getting hyperlinks from invalid slide."""
        pres = ppt.create_presentation()
        ppt.add_slide(pres, layout="Blank")

        links = ppt.get_hyperlinks(pres, 99)

        assert isinstance(links, list)
        assert len(links) == 0


class TestHyperlinkActions:
    """Tests for hyperlink navigation actions."""

    def test_action_first_slide(self):
        """Test first slide action."""
        pres = ppt.create_presentation()
        ppt.add_slide(pres, layout="Blank")
        ppt.add_slide(pres, layout="Blank")
        shape_id = ppt.add_shape(
            pres, 2, "rectangle",
            left="1in", top="1in", width="2in", height="1in"
        )

        result = ppt.add_hyperlink(pres, 2, shape_id, action="first_slide")

        assert result is True

    def test_action_last_slide(self):
        """Test last slide action."""
        pres = ppt.create_presentation()
        ppt.add_slide(pres, layout="Blank")
        ppt.add_slide(pres, layout="Blank")
        shape_id = ppt.add_shape(
            pres, 1, "rectangle",
            left="1in", top="1in", width="2in", height="1in"
        )

        result = ppt.add_hyperlink(pres, 1, shape_id, action="last_slide")

        assert result is True

    def test_action_end_show(self):
        """Test end show action."""
        pres = ppt.create_presentation()
        ppt.add_slide(pres, layout="Blank")
        shape_id = ppt.add_shape(
            pres, 1, "rectangle",
            left="1in", top="1in", width="2in", height="1in"
        )

        result = ppt.add_hyperlink(pres, 1, shape_id, action="end_show")

        assert result is True

    def test_invalid_action(self):
        """Test invalid action returns False."""
        pres = ppt.create_presentation()
        ppt.add_slide(pres, layout="Blank")
        shape_id = ppt.add_shape(
            pres, 1, "rectangle",
            left="1in", top="1in", width="2in", height="1in"
        )

        result = ppt.add_hyperlink(pres, 1, shape_id, action="invalid_action")

        assert result is False
