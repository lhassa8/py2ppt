"""Style guide rule definitions."""

from __future__ import annotations

from dataclasses import dataclass, field
from typing import Any


@dataclass
class StyleGuide:
    """Style guide rules for presentation validation.

    Defines constraints and requirements that presentations
    should follow for brand/style compliance.
    """

    # Content limits
    max_slides: int | None = None
    max_bullet_points: int | None = None  # Per slide
    max_words_per_bullet: int | None = None
    max_title_length: int | None = None  # Characters

    # Required elements
    required_layouts: list[str] = field(default_factory=list)
    required_first_slide: str | None = None  # Layout name

    # Typography
    allowed_fonts: set[str] | None = None
    forbidden_fonts: set[str] = field(default_factory=set)
    min_font_size: int | None = None  # Points
    max_font_size: int | None = None

    # Colors
    allowed_colors: set[str] | None = None  # Hex or token names
    forbidden_colors: set[str] = field(default_factory=set)

    # Custom rules
    custom_rules: dict[str, Any] = field(default_factory=dict)


def create_style_guide(spec: dict[str, Any]) -> StyleGuide:
    """Create a style guide from a specification dict.

    Args:
        spec: Style guide specification

    Returns:
        StyleGuide object

    Example:
        >>> rules = create_style_guide({
        ...     "max_bullet_points": 6,
        ...     "max_words_per_bullet": 12,
        ...     "max_slides": 30,
        ...     "required_layouts": ["title", "agenda"],
        ...     "forbidden_fonts": ["Comic Sans MS", "Papyrus"],
        ...     "min_font_size": 12,
        ... })
    """
    guide = StyleGuide()

    # Content limits
    guide.max_slides = spec.get("max_slides")
    guide.max_bullet_points = spec.get("max_bullet_points")
    guide.max_words_per_bullet = spec.get("max_words_per_bullet")
    guide.max_title_length = spec.get("max_title_length")

    # Required elements
    guide.required_layouts = spec.get("required_layouts", [])
    guide.required_first_slide = spec.get("required_first_slide")

    # Typography
    if "allowed_fonts" in spec:
        guide.allowed_fonts = set(spec["allowed_fonts"])
    if "forbidden_fonts" in spec:
        guide.forbidden_fonts = set(spec["forbidden_fonts"])

    min_size = spec.get("min_font_size")
    if min_size:
        if isinstance(min_size, str) and min_size.endswith("pt"):
            min_size = int(min_size[:-2])
        guide.min_font_size = int(min_size)

    max_size = spec.get("max_font_size")
    if max_size:
        if isinstance(max_size, str) and max_size.endswith("pt"):
            max_size = int(max_size[:-2])
        guide.max_font_size = int(max_size)

    # Colors
    if "allowed_colors" in spec:
        guide.allowed_colors = set(spec["allowed_colors"])
    if "forbidden_colors" in spec:
        guide.forbidden_colors = set(spec["forbidden_colors"])

    # Custom rules
    guide.custom_rules = spec.get("custom_rules", {})

    return guide
