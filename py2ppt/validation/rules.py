"""Style guide rule definitions."""

from __future__ import annotations

from collections.abc import Callable
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
    min_slide_count: int | None = None  # Minimum slides required

    # Required elements
    required_layouts: list[str] = field(default_factory=list)
    required_first_slide: str | None = None  # Layout name
    required_last_slide: str | None = None  # Layout name (e.g., "Thank You")

    # Typography
    allowed_fonts: set[str] | None = None
    forbidden_fonts: set[str] = field(default_factory=set)
    min_font_size: int | None = None  # Points
    max_font_size: int | None = None

    # Colors
    allowed_colors: set[str] | None = None  # Hex or token names
    forbidden_colors: set[str] = field(default_factory=set)
    require_theme_colors: bool = False  # Only allow theme colors

    # Images
    require_image_alt_text: bool = False  # Accessibility: require alt text
    max_images_per_slide: int | None = None
    min_image_resolution: tuple[int, int] | None = None  # (width, height) in pixels

    # Charts
    max_charts_per_slide: int | None = None
    require_chart_title: bool = False
    require_data_labels: bool = False

    # Tables
    max_table_rows: int | None = None
    max_table_columns: int | None = None
    require_table_header: bool = False

    # Accessibility
    check_color_contrast: bool = False  # Check text/background contrast
    min_contrast_ratio: float = 4.5  # WCAG AA standard
    require_slide_titles: bool = False  # Every slide should have a title

    # Output quality
    check_text_overflow: bool = False  # Check for text that might overflow
    check_empty_placeholders: bool = False  # Warn about empty placeholders

    # Custom rules (callable validators)
    custom_rules: dict[str, Any] = field(default_factory=dict)
    custom_validators: list[Callable] = field(default_factory=list)


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
        ...     "require_image_alt_text": True,
        ...     "check_color_contrast": True,
        ... })
    """
    guide = StyleGuide()

    # Content limits
    guide.max_slides = spec.get("max_slides")
    guide.min_slide_count = spec.get("min_slide_count")
    guide.max_bullet_points = spec.get("max_bullet_points")
    guide.max_words_per_bullet = spec.get("max_words_per_bullet")
    guide.max_title_length = spec.get("max_title_length")

    # Required elements
    guide.required_layouts = spec.get("required_layouts", [])
    guide.required_first_slide = spec.get("required_first_slide")
    guide.required_last_slide = spec.get("required_last_slide")

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
    guide.require_theme_colors = spec.get("require_theme_colors", False)

    # Images
    guide.require_image_alt_text = spec.get("require_image_alt_text", False)
    guide.max_images_per_slide = spec.get("max_images_per_slide")
    if "min_image_resolution" in spec:
        res = spec["min_image_resolution"]
        if isinstance(res, (list, tuple)) and len(res) == 2:
            guide.min_image_resolution = tuple(res)

    # Charts
    guide.max_charts_per_slide = spec.get("max_charts_per_slide")
    guide.require_chart_title = spec.get("require_chart_title", False)
    guide.require_data_labels = spec.get("require_data_labels", False)

    # Tables
    guide.max_table_rows = spec.get("max_table_rows")
    guide.max_table_columns = spec.get("max_table_columns")
    guide.require_table_header = spec.get("require_table_header", False)

    # Accessibility
    guide.check_color_contrast = spec.get("check_color_contrast", False)
    guide.min_contrast_ratio = spec.get("min_contrast_ratio", 4.5)
    guide.require_slide_titles = spec.get("require_slide_titles", False)

    # Output quality
    guide.check_text_overflow = spec.get("check_text_overflow", False)
    guide.check_empty_placeholders = spec.get("check_empty_placeholders", False)

    # Custom rules
    guide.custom_rules = spec.get("custom_rules", {})
    guide.custom_validators = spec.get("custom_validators", [])

    return guide


# Predefined style guides for common use cases

def corporate_style_guide() -> StyleGuide:
    """Create a typical corporate presentation style guide.

    Returns:
        StyleGuide configured for professional corporate presentations
    """
    return create_style_guide({
        "max_slides": 30,
        "max_bullet_points": 6,
        "max_words_per_bullet": 12,
        "max_title_length": 50,
        "min_font_size": 12,
        "forbidden_fonts": {"Comic Sans MS", "Papyrus", "Curlz MT"},
        "require_slide_titles": True,
        "require_image_alt_text": True,
        "require_theme_colors": True,
    })


def accessible_style_guide() -> StyleGuide:
    """Create an accessibility-focused style guide.

    Returns:
        StyleGuide configured for WCAG compliance
    """
    return create_style_guide({
        "min_font_size": 18,
        "check_color_contrast": True,
        "min_contrast_ratio": 4.5,
        "require_image_alt_text": True,
        "require_slide_titles": True,
        "max_words_per_bullet": 15,
    })


def minimal_style_guide() -> StyleGuide:
    """Create a minimal style guide with basic checks.

    Returns:
        StyleGuide with only essential rules
    """
    return create_style_guide({
        "max_bullet_points": 8,
        "max_slides": 50,
        "forbidden_fonts": {"Comic Sans MS"},
    })
