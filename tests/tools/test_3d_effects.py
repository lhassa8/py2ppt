"""Tests for 3D effect functionality."""

import tempfile
from pathlib import Path

import py2ppt as ppt


class TestAdd3DRotation:
    """Tests for add_3d_rotation function."""

    def test_add_rotation_default(self):
        """Test adding 3D rotation with defaults."""
        pres = ppt.create_presentation()
        ppt.add_slide(pres, layout="Blank")
        shape_id = ppt.add_shape(pres, 1, "rectangle", "1in", "1in", "2in", "2in")

        result = ppt.add_3d_rotation(pres, 1, shape_id)

        assert result is True

    def test_add_rotation_with_angles(self):
        """Test adding 3D rotation with custom angles."""
        pres = ppt.create_presentation()
        ppt.add_slide(pres, layout="Blank")
        shape_id = ppt.add_shape(pres, 1, "rectangle", "1in", "1in", "2in", "2in")

        result = ppt.add_3d_rotation(pres, 1, shape_id, lat=20, lon=30, rev=10)

        assert result is True

    def test_add_rotation_with_preset(self):
        """Test adding 3D rotation with camera preset."""
        pres = ppt.create_presentation()
        ppt.add_slide(pres, layout="Blank")
        shape_id = ppt.add_shape(pres, 1, "rectangle", "1in", "1in", "2in", "2in")

        result = ppt.add_3d_rotation(pres, 1, shape_id, preset="isometricLeftDown")

        assert result is True

    def test_add_rotation_perspective_preset(self):
        """Test adding 3D rotation with perspective preset."""
        pres = ppt.create_presentation()
        ppt.add_slide(pres, layout="Blank")
        shape_id = ppt.add_shape(pres, 1, "rectangle", "1in", "1in", "2in", "2in")

        result = ppt.add_3d_rotation(pres, 1, shape_id, preset="perspectiveLeft")

        assert result is True

    def test_add_rotation_invalid_shape(self):
        """Test adding rotation to nonexistent shape."""
        pres = ppt.create_presentation()
        ppt.add_slide(pres, layout="Blank")

        result = ppt.add_3d_rotation(pres, 1, 9999)

        assert result is False


class TestAdd3DDepth:
    """Tests for add_3d_depth function."""

    def test_add_depth_default(self):
        """Test adding 3D depth with default values."""
        pres = ppt.create_presentation()
        ppt.add_slide(pres, layout="Blank")
        shape_id = ppt.add_shape(pres, 1, "rectangle", "1in", "1in", "2in", "2in")

        result = ppt.add_3d_depth(pres, 1, shape_id)

        assert result is True

    def test_add_depth_with_points(self):
        """Test adding 3D depth with point value."""
        pres = ppt.create_presentation()
        ppt.add_slide(pres, layout="Blank")
        shape_id = ppt.add_shape(pres, 1, "rectangle", "1in", "1in", "2in", "2in")

        result = ppt.add_3d_depth(pres, 1, shape_id, "20pt")

        assert result is True

    def test_add_depth_with_color(self):
        """Test adding 3D depth with custom color."""
        pres = ppt.create_presentation()
        ppt.add_slide(pres, layout="Blank")
        shape_id = ppt.add_shape(pres, 1, "rectangle", "1in", "1in", "2in", "2in")

        result = ppt.add_3d_depth(pres, 1, shape_id, "15pt", color="#0066CC")

        assert result is True

    def test_add_depth_with_contour(self):
        """Test adding 3D depth with contour."""
        pres = ppt.create_presentation()
        ppt.add_slide(pres, layout="Blank")
        shape_id = ppt.add_shape(pres, 1, "rectangle", "1in", "1in", "2in", "2in")

        result = ppt.add_3d_depth(
            pres, 1, shape_id, "10pt",
            contour_width="2pt", contour_color="#000000"
        )

        assert result is True

    def test_add_depth_emu_value(self):
        """Test adding 3D depth with EMU value."""
        pres = ppt.create_presentation()
        ppt.add_slide(pres, layout="Blank")
        shape_id = ppt.add_shape(pres, 1, "rectangle", "1in", "1in", "2in", "2in")

        result = ppt.add_3d_depth(pres, 1, shape_id, 127000)  # ~10pt

        assert result is True

    def test_add_depth_invalid_shape(self):
        """Test adding depth to nonexistent shape."""
        pres = ppt.create_presentation()
        ppt.add_slide(pres, layout="Blank")

        result = ppt.add_3d_depth(pres, 1, 9999)

        assert result is False


class TestAddBevel:
    """Tests for add_bevel function."""

    def test_add_bevel_default(self):
        """Test adding bevel with default values."""
        pres = ppt.create_presentation()
        ppt.add_slide(pres, layout="Blank")
        shape_id = ppt.add_shape(pres, 1, "rectangle", "1in", "1in", "2in", "2in")

        result = ppt.add_bevel(pres, 1, shape_id)

        assert result is True

    def test_add_bevel_circle_type(self):
        """Test adding circle bevel."""
        pres = ppt.create_presentation()
        ppt.add_slide(pres, layout="Blank")
        shape_id = ppt.add_shape(pres, 1, "rectangle", "1in", "1in", "2in", "2in")

        result = ppt.add_bevel(pres, 1, shape_id, bevel_type="circle")

        assert result is True

    def test_add_bevel_soft_round(self):
        """Test adding soft round bevel."""
        pres = ppt.create_presentation()
        ppt.add_slide(pres, layout="Blank")
        shape_id = ppt.add_shape(pres, 1, "rectangle", "1in", "1in", "2in", "2in")

        result = ppt.add_bevel(pres, 1, shape_id, bevel_type="softRound")

        assert result is True

    def test_add_bevel_custom_size(self):
        """Test adding bevel with custom size."""
        pres = ppt.create_presentation()
        ppt.add_slide(pres, layout="Blank")
        shape_id = ppt.add_shape(pres, 1, "rectangle", "1in", "1in", "2in", "2in")

        result = ppt.add_bevel(pres, 1, shape_id, width="8pt", height="8pt")

        assert result is True

    def test_add_bevel_top_only(self):
        """Test adding bevel to top only."""
        pres = ppt.create_presentation()
        ppt.add_slide(pres, layout="Blank")
        shape_id = ppt.add_shape(pres, 1, "rectangle", "1in", "1in", "2in", "2in")

        result = ppt.add_bevel(pres, 1, shape_id, top=True, bottom=False)

        assert result is True

    def test_add_bevel_both_sides(self):
        """Test adding bevel to both top and bottom."""
        pres = ppt.create_presentation()
        ppt.add_slide(pres, layout="Blank")
        shape_id = ppt.add_shape(pres, 1, "rectangle", "1in", "1in", "2in", "2in")

        result = ppt.add_bevel(pres, 1, shape_id, top=True, bottom=True)

        assert result is True

    def test_add_bevel_invalid_shape(self):
        """Test adding bevel to nonexistent shape."""
        pres = ppt.create_presentation()
        ppt.add_slide(pres, layout="Blank")

        result = ppt.add_bevel(pres, 1, 9999)

        assert result is False


class TestRemove3DEffects:
    """Tests for remove_3d_effects function."""

    def test_remove_3d_effects(self):
        """Test removing 3D effects from a shape."""
        pres = ppt.create_presentation()
        ppt.add_slide(pres, layout="Blank")
        shape_id = ppt.add_shape(pres, 1, "rectangle", "1in", "1in", "2in", "2in")

        # Add some 3D effects
        ppt.add_3d_rotation(pres, 1, shape_id, preset="isometricLeftDown")
        ppt.add_3d_depth(pres, 1, shape_id, "10pt")

        result = ppt.remove_3d_effects(pres, 1, shape_id)

        assert result is True

    def test_remove_3d_effects_no_effects(self):
        """Test removing 3D effects when none exist."""
        pres = ppt.create_presentation()
        ppt.add_slide(pres, layout="Blank")
        shape_id = ppt.add_shape(pres, 1, "rectangle", "1in", "1in", "2in", "2in")

        result = ppt.remove_3d_effects(pres, 1, shape_id)

        assert result is True

    def test_remove_3d_effects_invalid_shape(self):
        """Test removing 3D effects from nonexistent shape."""
        pres = ppt.create_presentation()
        ppt.add_slide(pres, layout="Blank")

        result = ppt.remove_3d_effects(pres, 1, 9999)

        assert result is False


class TestGet3DPresets:
    """Tests for get_3d_presets function."""

    def test_get_presets_returns_list(self):
        """Test that get_3d_presets returns a list."""
        presets = ppt.get_3d_presets()

        assert isinstance(presets, list)
        assert len(presets) > 0

    def test_presets_contain_isometric(self):
        """Test that presets contain isometric options."""
        presets = ppt.get_3d_presets()

        isometric_presets = [p for p in presets if "isometric" in p.lower()]
        assert len(isometric_presets) > 0

    def test_presets_contain_perspective(self):
        """Test that presets contain perspective options."""
        presets = ppt.get_3d_presets()

        perspective_presets = [p for p in presets if "perspective" in p.lower()]
        assert len(perspective_presets) > 0

    def test_presets_contain_orthographic(self):
        """Test that presets contain orthographic options."""
        presets = ppt.get_3d_presets()

        assert "orthographicFront" in presets


class Test3DEffectsIntegration:
    """Integration tests for 3D effects."""

    def test_combine_rotation_and_depth(self):
        """Test combining rotation and depth effects."""
        pres = ppt.create_presentation()
        ppt.add_slide(pres, layout="Blank")
        shape_id = ppt.add_shape(pres, 1, "rectangle", "1in", "1in", "2in", "2in")

        # Add rotation
        result1 = ppt.add_3d_rotation(pres, 1, shape_id, preset="isometricLeftDown")
        # Add depth
        result2 = ppt.add_3d_depth(pres, 1, shape_id, "15pt", color="#003366")

        assert result1 is True
        assert result2 is True

    def test_combine_all_3d_effects(self):
        """Test combining rotation, depth, and bevel."""
        pres = ppt.create_presentation()
        ppt.add_slide(pres, layout="Blank")
        shape_id = ppt.add_shape(pres, 1, "rectangle", "1in", "1in", "2in", "2in")

        # Add all effects
        ppt.add_3d_rotation(pres, 1, shape_id, preset="perspectiveFront")
        ppt.add_3d_depth(pres, 1, shape_id, "12pt")
        ppt.add_bevel(pres, 1, shape_id, bevel_type="circle", top=True, bottom=True)

        # Verify presentation can be saved and loaded
        with tempfile.NamedTemporaryFile(suffix=".pptx", delete=False) as f:
            temp_pptx = f.name

        try:
            ppt.save_presentation(pres, temp_pptx)
            pres2 = ppt.open_presentation(temp_pptx)
            assert ppt.get_slide_count(pres2) == 1
        finally:
            Path(temp_pptx).unlink()

    def test_3d_effects_on_different_shapes(self):
        """Test 3D effects on various shape types."""
        pres = ppt.create_presentation()
        ppt.add_slide(pres, layout="Blank")

        # Create different shapes
        rect_id = ppt.add_shape(pres, 1, "rectangle", "1in", "1in", "1in", "1in")
        ellipse_id = ppt.add_shape(pres, 1, "ellipse", "3in", "1in", "1in", "1in")
        arrow_id = ppt.add_shape(pres, 1, "arrow_right", "1in", "3in", "2in", "1in")

        # Apply 3D effects to each
        assert ppt.add_3d_rotation(pres, 1, rect_id, preset="isometricLeftDown") is True
        assert ppt.add_bevel(pres, 1, ellipse_id, bevel_type="softRound") is True
        assert ppt.add_3d_depth(pres, 1, arrow_id, "8pt") is True

    def test_replace_3d_effects(self):
        """Test replacing 3D effects."""
        pres = ppt.create_presentation()
        ppt.add_slide(pres, layout="Blank")
        shape_id = ppt.add_shape(pres, 1, "rectangle", "1in", "1in", "2in", "2in")

        # Add initial effects
        ppt.add_3d_rotation(pres, 1, shape_id, preset="isometricLeftDown")

        # Replace with different rotation
        result = ppt.add_3d_rotation(pres, 1, shape_id, preset="perspectiveFront")

        assert result is True
