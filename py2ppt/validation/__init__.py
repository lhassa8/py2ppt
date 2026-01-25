"""Style guide validation for presentations."""

from .rules import (
    StyleGuide,
    accessible_style_guide,
    corporate_style_guide,
    create_style_guide,
    minimal_style_guide,
)
from .validator import (
    Severity,
    ValidationIssue,
    get_validation_summary,
    validate,
)

__all__ = [
    # Style guide
    "StyleGuide",
    "create_style_guide",
    # Predefined guides
    "corporate_style_guide",
    "accessible_style_guide",
    "minimal_style_guide",
    # Validation
    "validate",
    "ValidationIssue",
    "Severity",
    "get_validation_summary",
]
