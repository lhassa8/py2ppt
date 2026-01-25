"""Presentation validation engine."""

from __future__ import annotations

from dataclasses import dataclass
from enum import Enum

from ..core.presentation import Presentation
from ..oxml.shapes import Chart, Picture, Shape, Table
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

    # Presentation-level checks
    issues.extend(_validate_presentation_level(presentation, style_guide))

    # Check each slide
    for slide_num in range(1, presentation.slide_count + 1):
        slide = presentation.get_slide(slide_num)
        slide_issues = _validate_slide(slide, slide_num, style_guide)
        issues.extend(slide_issues)

    # Run custom validators
    for validator in style_guide.custom_validators:
        try:
            custom_issues = validator(presentation, style_guide)
            if custom_issues:
                issues.extend(custom_issues)
        except Exception as e:
            issues.append(
                ValidationIssue(
                    slide=None,
                    severity="warning",
                    rule="custom_validator_error",
                    message=f"Custom validator failed: {e}",
                )
            )

    return issues


def _validate_presentation_level(
    presentation: Presentation,
    guide: StyleGuide,
) -> list[ValidationIssue]:
    """Validate presentation-level rules."""
    issues = []

    # Check slide count limits
    if guide.max_slides and presentation.slide_count > guide.max_slides:
        issues.append(
            ValidationIssue(
                slide=None,
                severity="error",
                rule="max_slides",
                message=f"Too many slides ({presentation.slide_count} > {guide.max_slides} max)",
                details={
                    "actual": presentation.slide_count,
                    "max": guide.max_slides,
                },
            )
        )

    if guide.min_slide_count and presentation.slide_count < guide.min_slide_count:
        issues.append(
            ValidationIssue(
                slide=None,
                severity="warning",
                rule="min_slide_count",
                message=f"Too few slides ({presentation.slide_count} < {guide.min_slide_count} min)",
                details={
                    "actual": presentation.slide_count,
                    "min": guide.min_slide_count,
                },
            )
        )

    # Check required layouts
    if guide.required_layouts:
        layout_names = {name.lower() for name in presentation.get_layout_names()}
        for required in guide.required_layouts:
            if required.lower() not in layout_names:
                issues.append(
                    ValidationIssue(
                        slide=None,
                        severity="warning",
                        rule="required_layout",
                        message=f"Required layout '{required}' not available in template",
                    )
                )

    return issues


def _validate_slide(
    slide,
    slide_num: int,
    guide: StyleGuide,
) -> list[ValidationIssue]:
    """Validate a single slide."""
    issues = []

    # Content checks
    issues.extend(_check_content(slide, slide_num, guide))

    # Typography checks
    issues.extend(_check_typography(slide, slide_num, guide))

    # Image checks
    issues.extend(_check_images(slide, slide_num, guide))

    # Chart checks
    issues.extend(_check_charts(slide, slide_num, guide))

    # Table checks
    issues.extend(_check_tables(slide, slide_num, guide))

    # Accessibility checks
    issues.extend(_check_accessibility(slide, slide_num, guide))

    return issues


def _check_content(
    slide,
    slide_num: int,
    guide: StyleGuide,
) -> list[ValidationIssue]:
    """Check content-related rules."""
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

    # Check empty placeholders
    if guide.check_empty_placeholders:
        placeholders = slide.get_placeholders()
        for ph_type, shape in placeholders.items():
            if shape.text_frame and not shape.text_frame.text.strip():
                issues.append(
                    ValidationIssue(
                        slide=slide_num,
                        severity="info",
                        rule="empty_placeholder",
                        message=f"Empty placeholder: {ph_type}",
                        details={"placeholder_type": ph_type},
                    )
                )

    return issues


def _check_typography(
    slide,
    slide_num: int,
    guide: StyleGuide,
) -> list[ValidationIssue]:
    """Check typography rules."""
    issues = []

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


def _check_images(
    slide,
    slide_num: int,
    guide: StyleGuide,
) -> list[ValidationIssue]:
    """Check image-related rules."""
    issues = []

    # Count images
    images = [s for s in slide.shapes if isinstance(s, Picture)]

    # Check image count
    if guide.max_images_per_slide and len(images) > guide.max_images_per_slide:
        issues.append(
            ValidationIssue(
                slide=slide_num,
                severity="warning",
                rule="max_images_per_slide",
                message=f"Too many images ({len(images)} > {guide.max_images_per_slide} max)",
                details={
                    "actual": len(images),
                    "max": guide.max_images_per_slide,
                },
            )
        )

    # Check alt text
    if guide.require_image_alt_text:
        for i, img in enumerate(images):
            # Check if image has alt text (description)
            has_alt = hasattr(img, "alt_text") and img.alt_text
            if not has_alt:
                issues.append(
                    ValidationIssue(
                        slide=slide_num,
                        severity="warning",
                        rule="image_alt_text",
                        message=f"Image {i + 1} missing alt text (accessibility)",
                        details={"image_index": i},
                    )
                )

    return issues


def _check_charts(
    slide,
    slide_num: int,
    guide: StyleGuide,
) -> list[ValidationIssue]:
    """Check chart-related rules."""
    issues = []

    # Count charts
    charts = [s for s in slide.shapes if isinstance(s, Chart)]

    # Check chart count
    if guide.max_charts_per_slide and len(charts) > guide.max_charts_per_slide:
        issues.append(
            ValidationIssue(
                slide=slide_num,
                severity="warning",
                rule="max_charts_per_slide",
                message=f"Too many charts ({len(charts)} > {guide.max_charts_per_slide} max)",
                details={
                    "actual": len(charts),
                    "max": guide.max_charts_per_slide,
                },
            )
        )

    return issues


def _check_tables(
    slide,
    slide_num: int,
    guide: StyleGuide,
) -> list[ValidationIssue]:
    """Check table-related rules."""
    issues = []

    # Find tables
    tables = [s for s in slide.shapes if isinstance(s, Table)]

    for i, table in enumerate(tables):
        # Check row count
        if guide.max_table_rows:
            row_count = len(table.cells) if hasattr(table, "cells") else 0
            if row_count > guide.max_table_rows:
                issues.append(
                    ValidationIssue(
                        slide=slide_num,
                        severity="warning",
                        rule="max_table_rows",
                        message=f"Table {i + 1} has too many rows ({row_count} > {guide.max_table_rows} max)",
                        details={
                            "table_index": i,
                            "actual": row_count,
                            "max": guide.max_table_rows,
                        },
                    )
                )

    return issues


def _check_accessibility(
    slide,
    slide_num: int,
    guide: StyleGuide,
) -> list[ValidationIssue]:
    """Check accessibility-related rules."""
    issues = []

    # Check for slide title
    if guide.require_slide_titles:
        title = slide.get_title()
        if not title or not title.strip():
            issues.append(
                ValidationIssue(
                    slide=slide_num,
                    severity="warning",
                    rule="require_slide_title",
                    message="Slide is missing a title (accessibility)",
                )
            )

    # Check color contrast (basic implementation)
    if guide.check_color_contrast:
        contrast_issues = _check_color_contrast(slide, slide_num, guide.min_contrast_ratio)
        issues.extend(contrast_issues)

    return issues


def _check_color_contrast(
    slide,
    slide_num: int,
    min_ratio: float,
) -> list[ValidationIssue]:
    """Check text/background color contrast.

    Note: This is a simplified implementation. Full contrast checking
    would require parsing all color values and computing WCAG ratios.
    """
    issues = []

    # For now, just flag potential issues where custom colors are used
    # A full implementation would compute actual contrast ratios
    for shape in slide.shapes:
        if isinstance(shape, Shape) and shape.text_frame:
            # Check if text uses very light colors on light background
            # or very dark colors on dark background
            # This is a placeholder for more sophisticated contrast analysis
            pass

    return issues


def _get_fonts_used(slide) -> set:
    """Extract all fonts used in a slide."""
    fonts = set()

    for shape in slide.shapes:
        if isinstance(shape, Shape) and shape.text_frame:
            for para in shape.text_frame.paragraphs:
                for run in para.runs:
                    if run.properties and run.properties.font_family:
                        fonts.add(run.properties.font_family)

    return fonts


def _get_font_sizes(slide) -> set:
    """Extract all font sizes (in points) used in a slide."""
    sizes = set()

    for shape in slide.shapes:
        if isinstance(shape, Shape) and shape.text_frame:
            for para in shape.text_frame.paragraphs:
                for run in para.runs:
                    if run.properties and run.properties.font_size:
                        # Convert from centipoints to points
                        sizes.add(run.properties.font_size // 100)

    return sizes


def get_validation_summary(issues: list[ValidationIssue]) -> dict:
    """Get a summary of validation issues.

    Args:
        issues: List of validation issues

    Returns:
        Dict with counts by severity and rule

    Example:
        >>> summary = get_validation_summary(issues)
        >>> print(f"Errors: {summary['error_count']}")
        >>> print(f"Warnings: {summary['warning_count']}")
    """
    summary = {
        "total": len(issues),
        "error_count": 0,
        "warning_count": 0,
        "info_count": 0,
        "by_rule": {},
        "by_slide": {},
    }

    for issue in issues:
        # Count by severity
        if issue.severity == "error":
            summary["error_count"] += 1
        elif issue.severity == "warning":
            summary["warning_count"] += 1
        else:
            summary["info_count"] += 1

        # Count by rule
        if issue.rule not in summary["by_rule"]:
            summary["by_rule"][issue.rule] = 0
        summary["by_rule"][issue.rule] += 1

        # Count by slide
        slide_key = str(issue.slide) if issue.slide else "presentation"
        if slide_key not in summary["by_slide"]:
            summary["by_slide"][slide_key] = 0
        summary["by_slide"][slide_key] += 1

    return summary
