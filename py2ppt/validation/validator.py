"""Presentation validation engine."""

from __future__ import annotations

from dataclasses import dataclass
from enum import Enum

from ..core.presentation import Presentation
from ..oxml.shapes import Shape
from .rules import StyleGuide


class Severity(Enum):
    """Issue severity levels."""

    ERROR = "error"
    WARNING = "warning"
    INFO = "info"


@dataclass
class ValidationIssue:
    """A single validation issue found in the presentation."""

    slide: int | None  # Slide number, or None for presentation-level
    severity: str  # "error", "warning", "info"
    rule: str  # Rule identifier
    message: str  # Human-readable message
    details: dict | None = None


def validate(
    presentation: Presentation,
    style_guide: StyleGuide,
) -> list[ValidationIssue]:
    """Validate a presentation against a style guide.

    Args:
        presentation: The presentation to validate
        style_guide: Style guide rules to check

    Returns:
        List of ValidationIssue objects for any violations

    Example:
        >>> issues = validate(pres, rules)
        >>> for issue in issues:
        ...     print(f"Slide {issue.slide}: {issue.message}")
    """
    issues: list[ValidationIssue] = []

    # Check slide count
    if style_guide.max_slides and presentation.slide_count > style_guide.max_slides:
        issues.append(
                ValidationIssue(
                    slide=None,
                    severity="error",
                    rule="max_slides",
                    message=f"Too many slides ({presentation.slide_count} > {style_guide.max_slides} max)",
                    details={
                        "actual": presentation.slide_count,
                        "max": style_guide.max_slides,
                    },
                )
            )

    # Check required layouts
    if style_guide.required_layouts:
        layout_names = {name.lower() for name in presentation.get_layout_names()}
        for required in style_guide.required_layouts:
            if required.lower() not in layout_names:
                # Check if any slide uses a layout matching this name
                found = False
                for _i in range(1, presentation.slide_count + 1):
                    # TODO: Check slide layout
                    pass

                if not found:
                    issues.append(
                        ValidationIssue(
                            slide=None,
                            severity="warning",
                            rule="required_layout",
                            message=f"Required layout '{required}' not used in presentation",
                        )
                    )

    # Check each slide
    for slide_num in range(1, presentation.slide_count + 1):
        slide = presentation.get_slide(slide_num)
        slide_issues = _validate_slide(slide, slide_num, style_guide)
        issues.extend(slide_issues)

    return issues


def _validate_slide(
    slide,
    slide_num: int,
    guide: StyleGuide,
) -> list[ValidationIssue]:
    """Validate a single slide."""
    issues = []

    # Get body content
    body_content = slide.get_body()

    # Check bullet count
    if guide.max_bullet_points and len(body_content) > guide.max_bullet_points:
        issues.append(
                ValidationIssue(
                    slide=slide_num,
                    severity="error",
                    rule="max_bullet_points",
                    message=f"Too many bullet points ({len(body_content)} > {guide.max_bullet_points} max)",
                    details={
                        "actual": len(body_content),
                        "max": guide.max_bullet_points,
                    },
                )
            )

    # Check words per bullet
    if guide.max_words_per_bullet:
        for i, bullet in enumerate(body_content):
            word_count = len(bullet.split())
            if word_count > guide.max_words_per_bullet:
                issues.append(
                    ValidationIssue(
                        slide=slide_num,
                        severity="warning",
                        rule="max_words_per_bullet",
                        message=f"Bullet {i + 1} has too many words ({word_count} > {guide.max_words_per_bullet} max)",
                        details={
                            "bullet_index": i,
                            "actual": word_count,
                            "max": guide.max_words_per_bullet,
                            "text": bullet[:50] + "..." if len(bullet) > 50 else bullet,
                        },
                    )
                )

    # Check title length
    if guide.max_title_length:
        title = slide.get_title()
        if title and len(title) > guide.max_title_length:
            issues.append(
                ValidationIssue(
                    slide=slide_num,
                    severity="warning",
                    rule="max_title_length",
                    message=f"Title too long ({len(title)} > {guide.max_title_length} chars)",
                    details={
                        "actual": len(title),
                        "max": guide.max_title_length,
                    },
                )
            )

    # Check fonts
    if guide.forbidden_fonts or guide.allowed_fonts:
        fonts_used = _get_fonts_used(slide)
        for font in fonts_used:
            if font in guide.forbidden_fonts:
                issues.append(
                    ValidationIssue(
                        slide=slide_num,
                        severity="error",
                        rule="forbidden_font",
                        message=f"Font '{font}' is not allowed",
                        details={"font": font},
                    )
                )
            elif guide.allowed_fonts and font not in guide.allowed_fonts:
                issues.append(
                    ValidationIssue(
                        slide=slide_num,
                        severity="warning",
                        rule="allowed_fonts",
                        message=f"Font '{font}' not in allowed list",
                        details={
                            "font": font,
                            "allowed": list(guide.allowed_fonts),
                        },
                    )
                )

    # Check font sizes
    if guide.min_font_size or guide.max_font_size:
        sizes_used = _get_font_sizes(slide)
        for size_pt in sizes_used:
            if guide.min_font_size and size_pt < guide.min_font_size:
                issues.append(
                    ValidationIssue(
                        slide=slide_num,
                        severity="warning",
                        rule="min_font_size",
                        message=f"Font size {size_pt}pt is below minimum {guide.min_font_size}pt",
                        details={
                            "actual": size_pt,
                            "min": guide.min_font_size,
                        },
                    )
                )
            if guide.max_font_size and size_pt > guide.max_font_size:
                issues.append(
                    ValidationIssue(
                        slide=slide_num,
                        severity="warning",
                        rule="max_font_size",
                        message=f"Font size {size_pt}pt exceeds maximum {guide.max_font_size}pt",
                        details={
                            "actual": size_pt,
                            "max": guide.max_font_size,
                        },
                    )
                )

    return issues


def _get_fonts_used(slide) -> set:
    """Extract all fonts used in a slide."""
    fonts = set()

    for shape in slide.shapes:
        if isinstance(shape, Shape) and shape.text_frame:
            for para in shape.text_frame.paragraphs:
                for run in para.runs:
                    if run.properties.font_family:
                        fonts.add(run.properties.font_family)

    return fonts


def _get_font_sizes(slide) -> set:
    """Extract all font sizes (in points) used in a slide."""
    sizes = set()

    for shape in slide.shapes:
        if isinstance(shape, Shape) and shape.text_frame:
            for para in shape.text_frame.paragraphs:
                for run in para.runs:
                    if run.properties.font_size:
                        # Convert from centipoints to points
                        sizes.add(run.properties.font_size // 100)

    return sizes
