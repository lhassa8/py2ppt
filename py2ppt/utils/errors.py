"""Error types and structured error handling for AI agents.

This module provides:
1. Exception classes for Python-style error handling
2. Structured result types for AI tool-calling interfaces
"""

from __future__ import annotations

from dataclasses import dataclass, field
from typing import Any, Generic, TypeVar

# === Exception Classes ===


class Py2PptError(Exception):
    """Base exception for py2ppt errors."""

    def __init__(
        self,
        message: str,
        code: str | None = None,
        details: dict[str, Any] | None = None,
    ) -> None:
        super().__init__(message)
        self.code = code or "UNKNOWN_ERROR"
        self.details = details or {}

    def to_dict(self) -> dict[str, Any]:
        """Convert to dictionary for structured error responses."""
        return {
            "code": self.code,
            "message": str(self),
            **self.details,
        }


class LayoutNotFoundError(Py2PptError):
    """Layout name not found in template."""

    def __init__(
        self,
        layout_name: str,
        available: list[str] | None = None,
        suggestion: str | None = None,
    ) -> None:
        message = f"Layout '{layout_name}' not found in template"
        details: dict[str, Any] = {"requested": layout_name}

        if available:
            details["available"] = available
            message += f". Available layouts: {', '.join(available)}"

        if suggestion:
            details["suggestion"] = suggestion
            details["did_you_mean"] = suggestion
            message += f". Did you mean '{suggestion}'?"

        super().__init__(message, code="LAYOUT_NOT_FOUND", details=details)


class SlideNotFoundError(Py2PptError):
    """Slide number out of range."""

    def __init__(
        self,
        slide_number: int,
        total_slides: int,
    ) -> None:
        message = f"Slide {slide_number} not found. Presentation has {total_slides} slides (1-{total_slides})"
        details = {
            "requested": slide_number,
            "total_slides": total_slides,
            "valid_range": f"1-{total_slides}" if total_slides > 0 else "none",
        }
        super().__init__(message, code="SLIDE_NOT_FOUND", details=details)


class PlaceholderNotFoundError(Py2PptError):
    """Placeholder not found in slide."""

    def __init__(
        self,
        placeholder: str,
        slide_number: int,
        available: list[str] | None = None,
    ) -> None:
        message = f"Placeholder '{placeholder}' not found in slide {slide_number}"
        details: dict[str, Any] = {
            "requested": placeholder,
            "slide_number": slide_number,
        }

        if available:
            details["available"] = available
            message += f". Available: {', '.join(available)}"

        super().__init__(message, code="PLACEHOLDER_NOT_FOUND", details=details)


class InvalidTemplateError(Py2PptError):
    """Template file is invalid or corrupted."""

    def __init__(self, message: str, path: str | None = None) -> None:
        details = {}
        if path:
            details["path"] = path
        super().__init__(message, code="INVALID_TEMPLATE", details=details)


class ContentError(Py2PptError):
    """Error with content (e.g., invalid image, malformed data)."""

    def __init__(self, message: str, content_type: str | None = None) -> None:
        details = {}
        if content_type:
            details["content_type"] = content_type
        super().__init__(message, code="CONTENT_ERROR", details=details)


class StyleError(Py2PptError):
    """Error with styling (e.g., invalid color, unknown font)."""

    def __init__(
        self,
        message: str,
        property_name: str | None = None,
        value: str | None = None,
    ) -> None:
        details: dict[str, Any] = {}
        if property_name:
            details["property"] = property_name
        if value:
            details["value"] = value
        super().__init__(message, code="STYLE_ERROR", details=details)


# === Structured Results for Tool Calling ===

T = TypeVar("T")


@dataclass
class ToolError:
    """Structured error for AI tool responses."""

    code: str
    message: str
    suggestion: str | None = None
    did_you_mean: str | None = None
    available: list[str] | None = None
    details: dict[str, Any] = field(default_factory=dict)


@dataclass
class ToolResult(Generic[T]):
    """Result wrapper for tool functions.

    This provides a consistent interface for AI agents to handle
    both success and error cases without exceptions.

    Example:
        result = add_slide(pres, layout="title")
        if result.success:
            slide_num = result.value
        else:
            print(result.error.message)
            print(result.error.available)  # AI can self-correct
    """

    success: bool
    value: T | None = None
    error: ToolError | None = None

    def __bool__(self) -> bool:
        """Allow using result in boolean context."""
        return self.success

    def unwrap(self) -> T:
        """Get the value, raising if error.

        Use this when you want exception-style error handling.
        """
        if not self.success:
            raise Py2PptError(
                self.error.message if self.error else "Unknown error",
                code=self.error.code if self.error else "UNKNOWN",
                details=self.error.details if self.error else {},
            )
        return self.value  # type: ignore

    def to_dict(self) -> dict[str, Any]:
        """Convert to dictionary for JSON serialization."""
        if self.success:
            return {"success": True, "value": self.value}
        else:
            error_dict = {
                "code": self.error.code if self.error else "UNKNOWN",
                "message": self.error.message if self.error else "Unknown error",
            }
            if self.error:
                if self.error.suggestion:
                    error_dict["suggestion"] = self.error.suggestion
                if self.error.did_you_mean:
                    error_dict["did_you_mean"] = self.error.did_you_mean
                if self.error.available:
                    error_dict["available"] = self.error.available
                if self.error.details:
                    error_dict.update(self.error.details)
            return {"success": False, "error": error_dict}


def success(value: T) -> ToolResult[T]:
    """Create a successful result."""
    return ToolResult(success=True, value=value)


def error(
    code: str,
    message: str,
    *,
    suggestion: str | None = None,
    did_you_mean: str | None = None,
    available: list[str] | None = None,
    **details: Any,
) -> ToolResult[Any]:
    """Create an error result.

    Args:
        code: Error code (e.g., "LAYOUT_NOT_FOUND")
        message: Human-readable error message
        suggestion: Suggested fix
        did_you_mean: Alternative that might be correct
        available: List of valid options
        **details: Additional error details
    """
    return ToolResult(
        success=False,
        error=ToolError(
            code=code,
            message=message,
            suggestion=suggestion,
            did_you_mean=did_you_mean,
            available=available,
            details=details,
        ),
    )


def from_exception(exc: Py2PptError) -> ToolResult[Any]:
    """Convert a Py2PptError exception to a ToolResult."""
    return ToolResult(
        success=False,
        error=ToolError(
            code=exc.code,
            message=str(exc),
            available=exc.details.get("available"),
            did_you_mean=exc.details.get("did_you_mean"),
            suggestion=exc.details.get("suggestion"),
            details={
                k: v
                for k, v in exc.details.items()
                if k not in ("available", "did_you_mean", "suggestion")
            },
        ),
    )


def find_similar(name: str, options: list[str], threshold: float = 0.6) -> str | None:
    """Find the most similar option to a given name.

    Uses a simple similarity measure (Jaccard index of character n-grams).

    Args:
        name: Name to match
        options: List of valid options
        threshold: Minimum similarity score (0.0 to 1.0)

    Returns:
        Most similar option or None if nothing is close enough
    """
    if not options:
        return None

    def ngrams(s: str, n: int = 2) -> set:
        s = s.lower()
        return {s[i : i + n] for i in range(max(1, len(s) - n + 1))}

    def jaccard(a: set, b: set) -> float:
        if not a and not b:
            return 1.0
        if not a or not b:
            return 0.0
        return len(a & b) / len(a | b)

    name_ngrams = ngrams(name)
    best_score = 0.0
    best_match = None

    for option in options:
        score = jaccard(name_ngrams, ngrams(option))
        if score > best_score:
            best_score = score
            best_match = option

    if best_score >= threshold:
        return best_match

    return None
