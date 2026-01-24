"""Style guide validation for presentations."""

from .rules import StyleGuide, create_style_guide
from .validator import ValidationIssue, validate

__all__ = [
    "StyleGuide",
    "create_style_guide",
    "validate",
    "ValidationIssue",
]
