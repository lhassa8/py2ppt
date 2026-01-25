"""Tests for shape effects functionality."""

import py2ppt as ppt


class TestAddShadow:
    """Tests for add_shadow function."""

    def test_add_shadow_default(self):
        """Test adding default shadow to shape."""
        pres = ppt.create_presentation()
        ppt.add_slide(pres, layout="Blank")
        shape_id = ppt.add_shape(
            pres, 1, "rectangle",
            left="1in", top="1in", width="2in", height="1in"
        )

        result = ppt.add_shadow(pres, 1, shape_id)

        assert result is True

    def test_add_shadow_custom(self):
        """Test adding shadow with custom parameters."""
        pres = ppt.create_presentation()
        ppt.add_slide(pres, layout="Blank")
        shape_id = ppt.add_shape(
            pres, 1, "rectangle",
            left="1in", top="1in", width="2in", height="1in"
        )

        result = ppt.add_shadow(
            pres, 1, shape_id,
            blur=6,
            distance=5,
            direction=90,
            color="#333333",
            transparency=50
        )

        assert result is True

    def test_add_inner_shadow(self):
        """Test adding inner shadow."""
        pres = ppt.create_presentation()
        ppt.add_slide(pres, layout="Blank")
        shape_id = ppt.add_shape(
            pres, 1, "rectangle",
            left="1in", top="1in", width="2in", height="1in"
        )

        result = ppt.add_shadow(pres, 1, shape_id, shadow_type="inner")

        assert result is True

    def test_add_shadow_nonexistent_shape(self):
        """Test adding shadow to non-existent shape."""
        pres = ppt.create_presentation()
        ppt.add_slide(pres, layout="Blank")

        result = ppt.add_shadow(pres, 1, 9999)

        assert result is False


class TestAddGlow:
    """Tests for add_glow function."""

    def test_add_glow_default(self):
        """Test adding default glow to shape."""
        pres = ppt.create_presentation()
        ppt.add_slide(pres, layout="Blank")
        shape_id = ppt.add_shape(
            pres, 1, "rectangle",
            left="1in", top="1in", width="2in", height="1in"
        )

        result = ppt.add_glow(pres, 1, shape_id)

        assert result is True

    def test_add_glow_custom(self):
        """Test adding glow with custom parameters."""
        pres = ppt.create_presentation()
        ppt.add_slide(pres, layout="Blank")
        shape_id = ppt.add_shape(
            pres, 1, "rectangle",
            left="1in", top="1in", width="2in", height="1in"
        )

        result = ppt.add_glow(
            pres, 1, shape_id,
            radius=10,
            color="#00FF00",
            transparency=30
        )

        assert result is True

    def test_add_glow_nonexistent_shape(self):
        """Test adding glow to non-existent shape."""
        pres = ppt.create_presentation()
        ppt.add_slide(pres, layout="Blank")

        result = ppt.add_glow(pres, 1, 9999)

        assert result is False


class TestAddReflection:
    """Tests for add_reflection function."""

    def test_add_reflection_default(self):
        """Test adding default reflection to shape."""
        pres = ppt.create_presentation()
        ppt.add_slide(pres, layout="Blank")
        shape_id = ppt.add_shape(
            pres, 1, "rectangle",
            left="1in", top="1in", width="2in", height="1in"
        )

        result = ppt.add_reflection(pres, 1, shape_id)

        assert result is True

    def test_add_reflection_custom(self):
        """Test adding reflection with custom parameters."""
        pres = ppt.create_presentation()
        ppt.add_slide(pres, layout="Blank")
        shape_id = ppt.add_shape(
            pres, 1, "rectangle",
            left="1in", top="1in", width="2in", height="1in"
        )

        result = ppt.add_reflection(
            pres, 1, shape_id,
            size=30,
            blur=1,
            start_transparency=10
        )

        assert result is True

    def test_add_reflection_nonexistent_shape(self):
        """Test adding reflection to non-existent shape."""
        pres = ppt.create_presentation()
        ppt.add_slide(pres, layout="Blank")

        result = ppt.add_reflection(pres, 1, 9999)

        assert result is False


class TestRemoveEffects:
    """Tests for remove_effects function."""

    def test_remove_effects(self):
        """Test removing effects from shape."""
        pres = ppt.create_presentation()
        ppt.add_slide(pres, layout="Blank")
        shape_id = ppt.add_shape(
            pres, 1, "rectangle",
            left="1in", top="1in", width="2in", height="1in"
        )

        # Add effects first
        ppt.add_shadow(pres, 1, shape_id)
        ppt.add_glow(pres, 1, shape_id)

        # Remove all effects
        result = ppt.remove_effects(pres, 1, shape_id)

        assert result is True

    def test_remove_effects_no_effects(self):
        """Test removing effects when none exist."""
        pres = ppt.create_presentation()
        ppt.add_slide(pres, layout="Blank")
        shape_id = ppt.add_shape(
            pres, 1, "rectangle",
            left="1in", top="1in", width="2in", height="1in"
        )

        # Should not raise even if no effects
        result = ppt.remove_effects(pres, 1, shape_id)

        assert result is True

    def test_remove_effects_nonexistent_shape(self):
        """Test removing effects from non-existent shape."""
        pres = ppt.create_presentation()
        ppt.add_slide(pres, layout="Blank")

        result = ppt.remove_effects(pres, 1, 9999)

        assert result is False
