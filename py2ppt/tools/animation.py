"""Animation tool functions.

Functions for adding slide transitions and shape animations.
"""

from __future__ import annotations

from typing import Literal

from ..core.presentation import Presentation
from ..oxml.animation import (
    ANIMATION_MAP,
    TRANSITION_MAP,
    AnimationSequence,
    ShapeAnimation,
    SlideTransition,
)
from ..oxml.ns import qn
from ..oxml.slide import update_slide_in_package


def set_slide_transition(
    presentation: Presentation,
    slide_number: int,
    transition: str = "fade",
    *,
    duration: int = 1000,
    advance_on_click: bool = True,
    advance_time: int | None = None,
    direction: str | None = None,
) -> None:
    """Set a transition effect for a slide.

    The transition plays when navigating to this slide.

    Args:
        presentation: The presentation to modify
        slide_number: Slide number (1-indexed)
        transition: Transition type. Options:
            - "fade": Fade in (default)
            - "push": Push previous slide out
            - "wipe": Wipe effect
            - "split": Split and reveal
            - "blinds": Venetian blinds effect
            - "checker": Checkerboard pattern
            - "circle": Circle reveal
            - "dissolve": Dissolve effect
            - "comb": Comb effect
            - "cover": Cover previous slide
            - "cut": Instant cut
            - "diamond": Diamond reveal
            - "plus": Plus/cross reveal
            - "random": Random transition
            - "strips": Diagonal strips
            - "wedge": Wedge reveal
            - "wheel": Wheel effect
            - "zoom": Zoom in/out
            - "none": No transition
        duration: Duration in milliseconds (default 1000)
        advance_on_click: Advance slide on mouse click (default True)
        advance_time: Auto-advance time in ms (None = manual only)
        direction: Direction hint (e.g., "l", "r", "u", "d")

    Example:
        >>> # Simple fade transition
        >>> set_slide_transition(pres, 1, "fade")

        >>> # Push with auto-advance after 5 seconds
        >>> set_slide_transition(pres, 2, "push",
        ...     advance_time=5000, direction="l")

        >>> # Fast wipe effect
        >>> set_slide_transition(pres, 3, "wipe", duration=500)
    """
    if transition.lower() not in TRANSITION_MAP:
        available = [k for k, v in TRANSITION_MAP.items() if v is not None]
        raise ValueError(
            f"Unknown transition: {transition}. "
            f"Available: {', '.join(available)}"
        )

    slide = presentation.get_slide(slide_number)
    slide_part = slide._part

    # Create transition element
    trans = SlideTransition(
        type=transition,
        duration=duration,
        advance_on_click=advance_on_click,
        advance_time=advance_time,
        direction=direction,
    )

    # Get slide element
    slide_elem = slide_part._element

    # Remove existing transition if present
    existing = slide_elem.find(qn("p:transition"))
    if existing is not None:
        slide_elem.remove(existing)

    # Add new transition (should be after p:cSld, before p:clrMapOvr)
    trans_xml = trans.to_xml()

    # Find insertion point
    csld = slide_elem.find(qn("p:cSld"))
    if csld is not None:
        csld_idx = list(slide_elem).index(csld)
        slide_elem.insert(csld_idx + 1, trans_xml)
    else:
        slide_elem.append(trans_xml)

    # Save changes
    update_slide_in_package(
        presentation._package,
        slide_number,
        slide_part,
    )


def add_animation(
    presentation: Presentation,
    slide_number: int,
    shape_id: int | str,
    effect: str = "appear",
    *,
    trigger: Literal["onClick", "withPrevious", "afterPrevious"] = "onClick",
    duration: int = 500,
    delay: int = 0,
) -> None:
    """Add an animation effect to a shape.

    Animations play in sequence during the slide show.

    Args:
        presentation: The presentation to modify
        slide_number: Slide number (1-indexed)
        shape_id: Shape ID (int) or name (str) to animate
        effect: Animation effect. Options include:
            Entrance effects:
            - "appear": Instantly appear (default)
            - "fade": Fade in
            - "fly_in": Fly in from edge
            - "float_in": Float up while fading
            - "split": Split and reveal
            - "wipe": Wipe effect
            - "zoom": Zoom in
            - "bounce": Bounce in
            Emphasis effects:
            - "pulse": Briefly enlarge
            - "spin": Rotate
            - "teeter": Rock back and forth
            - "grow_shrink": Grow then shrink
            Exit effects:
            - "disappear": Instantly vanish
            - "fade_out": Fade away
            - "fly_out": Fly off screen
            - "zoom_out": Zoom out
        trigger: When to play:
            - "onClick": On mouse click (default)
            - "withPrevious": With previous animation
            - "afterPrevious": After previous completes
        duration: Duration in milliseconds (default 500)
        delay: Delay before starting in milliseconds

    Example:
        >>> # Fade in title on click
        >>> add_animation(pres, 1, "Title", "fade")

        >>> # Make bullets appear one by one
        >>> add_animation(pres, 1, 3, "appear", trigger="onClick")
        >>> add_animation(pres, 1, 4, "appear", trigger="onClick")

        >>> # Emphasis effect with delay
        >>> add_animation(pres, 1, "Chart", "pulse",
        ...     trigger="afterPrevious", delay=500)
    """
    if effect.lower() not in ANIMATION_MAP:
        available = list(ANIMATION_MAP.keys())
        raise ValueError(
            f"Unknown animation effect: {effect}. "
            f"Available: {', '.join(available[:10])}..."
        )

    slide = presentation.get_slide(slide_number)
    slide_part = slide._part

    # Resolve shape_id if string (name)
    if isinstance(shape_id, str):
        actual_id = _find_shape_by_name(slide, shape_id)
        if actual_id is None:
            raise ValueError(f"Shape not found: {shape_id}")
        shape_id = actual_id

    # Get or create animation sequence
    slide_elem = slide_part._element
    timing_elem = slide_elem.find(qn("p:timing"))

    # Create animation
    anim = ShapeAnimation(
        shape_id=shape_id,
        effect=effect.lower(),
        trigger=trigger,
        duration=duration,
        delay=delay,
    )

    # For now, create new animation sequence
    # (Full implementation would parse existing and append)
    sequence = AnimationSequence(animations=[anim])

    # Remove existing timing
    if timing_elem is not None:
        slide_elem.remove(timing_elem)

    # Add new timing
    timing_xml = sequence.to_xml()
    if timing_xml is not None:
        slide_elem.append(timing_xml)

    # Save changes
    update_slide_in_package(
        presentation._package,
        slide_number,
        slide_part,
    )


def remove_animations(
    presentation: Presentation,
    slide_number: int,
) -> None:
    """Remove all animations from a slide.

    Args:
        presentation: The presentation to modify
        slide_number: Slide number (1-indexed)

    Example:
        >>> remove_animations(pres, 1)
    """
    slide = presentation.get_slide(slide_number)
    slide_part = slide._part
    slide_elem = slide_part._element

    # Remove timing element
    timing_elem = slide_elem.find(qn("p:timing"))
    if timing_elem is not None:
        slide_elem.remove(timing_elem)

    # Save changes
    update_slide_in_package(
        presentation._package,
        slide_number,
        slide_part,
    )


def remove_transition(
    presentation: Presentation,
    slide_number: int,
) -> None:
    """Remove transition from a slide.

    Args:
        presentation: The presentation to modify
        slide_number: Slide number (1-indexed)

    Example:
        >>> remove_transition(pres, 1)
    """
    slide = presentation.get_slide(slide_number)
    slide_part = slide._part
    slide_elem = slide_part._element

    # Remove transition element
    trans_elem = slide_elem.find(qn("p:transition"))
    if trans_elem is not None:
        slide_elem.remove(trans_elem)

    # Save changes
    update_slide_in_package(
        presentation._package,
        slide_number,
        slide_part,
    )


def get_available_transitions() -> list[str]:
    """Get list of available transition types.

    Returns:
        List of transition type names

    Example:
        >>> transitions = get_available_transitions()
        >>> print(transitions)
        ['fade', 'push', 'wipe', ...]
    """
    return [k for k, v in TRANSITION_MAP.items() if v is not None]


def get_available_animations() -> dict[str, list[str]]:
    """Get list of available animation effects by category.

    Returns:
        Dict mapping category to list of effect names

    Example:
        >>> anims = get_available_animations()
        >>> print(anims["entrance"])
        ['appear', 'fade', 'fly_in', ...]
    """
    categories: dict[str, list[str]] = {
        "entrance": [],
        "emphasis": [],
        "exit": [],
        "motion": [],
    }

    category_map = {
        "entr": "entrance",
        "emph": "emphasis",
        "exit": "exit",
        "path": "motion",
    }

    for name, info in ANIMATION_MAP.items():
        cat = category_map.get(info["preset_class"], "entrance")
        categories[cat].append(name)

    return categories


def _find_shape_by_name(slide, name: str) -> int | None:
    """Find shape ID by name."""
    for shape in slide.shapes:
        if hasattr(shape, "name") and shape.name == name:
            return shape.id
    return None
