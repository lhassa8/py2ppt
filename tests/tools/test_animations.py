"""Tests for animation functionality."""


import py2ppt as ppt


class TestSlideTransitions:
    """Tests for slide transition functions."""

    def test_set_fade_transition(self, tmp_path):
        """Test setting a fade transition."""
        pres = ppt.create_presentation()
        ppt.add_slide(pres, layout="Title Slide")

        ppt.set_slide_transition(pres, 1, transition="fade", duration=1000)

        path = tmp_path / "fade_transition.pptx"
        ppt.save_presentation(pres, str(path))
        assert path.exists()

    def test_set_push_transition(self, tmp_path):
        """Test setting a push transition."""
        pres = ppt.create_presentation()
        ppt.add_slide(pres, layout="Title Slide")

        ppt.set_slide_transition(
            pres, 1, transition="push", duration=500, direction="left"
        )

        path = tmp_path / "push_transition.pptx"
        ppt.save_presentation(pres, str(path))
        assert path.exists()

    def test_remove_transition(self, tmp_path):
        """Test removing a transition."""
        pres = ppt.create_presentation()
        ppt.add_slide(pres, layout="Title Slide")

        # Add then remove
        ppt.set_slide_transition(pres, 1, transition="fade")
        ppt.remove_transition(pres, 1)

        path = tmp_path / "no_transition.pptx"
        ppt.save_presentation(pres, str(path))
        assert path.exists()

    def test_get_available_transitions(self):
        """Test getting list of available transitions."""
        transitions = ppt.get_available_transitions()
        assert isinstance(transitions, list)
        assert "fade" in transitions
        assert "push" in transitions
        assert "wipe" in transitions


class TestShapeAnimations:
    """Tests for shape animation functions."""

    def test_add_appear_animation(self, tmp_path):
        """Test adding an appear animation."""
        pres = ppt.create_presentation()
        ppt.add_slide(pres, layout="Title and Content")
        ppt.set_title(pres, 1, "Animated Slide")

        # Get a shape to animate (title placeholder)
        slide = pres.get_slide(1)
        shapes = list(slide.shapes)
        if shapes:
            ppt.add_animation(pres, 1, shapes[0].id, effect="appear")

        path = tmp_path / "appear_animation.pptx"
        ppt.save_presentation(pres, str(path))
        assert path.exists()

    def test_add_fade_animation(self, tmp_path):
        """Test adding a fade animation."""
        pres = ppt.create_presentation()
        ppt.add_slide(pres, layout="Title Slide")
        ppt.set_title(pres, 1, "Fade In")

        slide = pres.get_slide(1)
        shapes = list(slide.shapes)
        if shapes:
            ppt.add_animation(
                pres, 1, shapes[0].id, effect="fade", duration=1000, delay=500
            )

        path = tmp_path / "fade_animation.pptx"
        ppt.save_presentation(pres, str(path))
        assert path.exists()

    def test_remove_animations(self, tmp_path):
        """Test removing all animations from a slide."""
        pres = ppt.create_presentation()
        ppt.add_slide(pres, layout="Title Slide")

        # Add and remove
        slide = pres.get_slide(1)
        shapes = list(slide.shapes)
        if shapes:
            ppt.add_animation(pres, 1, shapes[0].id, effect="appear")
        ppt.remove_animations(pres, 1)

        path = tmp_path / "no_animations.pptx"
        ppt.save_presentation(pres, str(path))
        assert path.exists()

    def test_get_available_animations(self):
        """Test getting list of available animations."""
        animations = ppt.get_available_animations()
        assert isinstance(animations, dict)
        assert "entrance" in animations
        assert "appear" in animations["entrance"]
