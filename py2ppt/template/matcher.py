"""Fuzzy layout and placeholder matching.

Provides intelligent matching between user-provided layout names
and actual layout names in templates.
"""

from __future__ import annotations

from dataclasses import dataclass
from typing import Dict, List, Optional, Tuple


# Common layout name patterns and their canonical forms
LAYOUT_PATTERNS: Dict[str, List[str]] = {
    "title slide": [
        "title slide", "title_slide", "titleslide",
        "title page", "title_page", "cover", "cover slide",
    ],
    "title and content": [
        "title and content", "title_and_content", "titleandcontent",
        "title content", "title_content", "content",
        "title & content", "title + content",
    ],
    "section header": [
        "section header", "section_header", "sectionheader",
        "section", "section slide", "divider",
    ],
    "two content": [
        "two content", "two_content", "twocontent",
        "two column", "two_column", "twocolumn",
        "2 content", "2 column", "dual content",
    ],
    "comparison": [
        "comparison", "compare", "side by side",
        "side_by_side", "versus", "vs",
    ],
    "title only": [
        "title only", "title_only", "titleonly",
        "title", "just title",
    ],
    "blank": [
        "blank", "empty", "none", "no content",
    ],
    "content with caption": [
        "content with caption", "content_with_caption",
        "caption", "captioned",
    ],
    "picture with caption": [
        "picture with caption", "picture_with_caption",
        "image with caption", "photo with caption",
    ],
}


def normalize_name(name: str) -> str:
    """Normalize a name for comparison.

    - Lowercase
    - Replace underscores and hyphens with spaces
    - Strip whitespace
    - Remove common prefixes/suffixes
    """
    name = name.lower().strip()
    name = name.replace("_", " ").replace("-", " ")
    name = " ".join(name.split())  # Normalize whitespace

    # Remove common prefixes
    prefixes = ["slide master ", "slide layout ", "layout "]
    for prefix in prefixes:
        if name.startswith(prefix):
            name = name[len(prefix):]

    return name


def find_canonical_layout(name: str) -> Optional[str]:
    """Find the canonical layout name for a given input.

    Args:
        name: User-provided layout name

    Returns:
        Canonical layout name or None if no match
    """
    normalized = normalize_name(name)

    # Direct match on canonical names
    if normalized in LAYOUT_PATTERNS:
        return normalized

    # Check all patterns
    for canonical, patterns in LAYOUT_PATTERNS.items():
        if normalized in patterns:
            return canonical
        # Partial match
        for pattern in patterns:
            if normalized in pattern or pattern in normalized:
                return canonical

    return None


@dataclass
class LayoutMatch:
    """Result of a layout matching operation."""

    found: bool
    layout_name: Optional[str] = None
    layout_index: Optional[int] = None
    confidence: float = 0.0
    alternatives: List[str] = None

    def __post_init__(self):
        if self.alternatives is None:
            self.alternatives = []


class LayoutMatcher:
    """Matches user-provided layout names to actual layouts in a presentation."""

    def __init__(self, layout_names: List[str]) -> None:
        """Initialize with available layout names.

        Args:
            layout_names: List of actual layout names from the presentation
        """
        self._layouts = layout_names
        self._normalized = {
            normalize_name(name): (name, i)
            for i, name in enumerate(layout_names)
        }

    def match(self, query: str) -> LayoutMatch:
        """Find the best matching layout for a query.

        Args:
            query: User-provided layout name

        Returns:
            LayoutMatch with results
        """
        normalized_query = normalize_name(query)

        # Exact match on normalized name
        if normalized_query in self._normalized:
            name, idx = self._normalized[normalized_query]
            return LayoutMatch(
                found=True,
                layout_name=name,
                layout_index=idx,
                confidence=1.0,
            )

        # Try canonical lookup
        canonical = find_canonical_layout(query)
        if canonical:
            # Find a layout matching this canonical name
            for norm_name, (name, idx) in self._normalized.items():
                if find_canonical_layout(name) == canonical:
                    return LayoutMatch(
                        found=True,
                        layout_name=name,
                        layout_index=idx,
                        confidence=0.9,
                    )

        # Fuzzy matching
        best_match = None
        best_score = 0.0

        for norm_name, (name, idx) in self._normalized.items():
            score = self._similarity(normalized_query, norm_name)
            if score > best_score:
                best_score = score
                best_match = (name, idx)

        if best_score >= 0.6:
            return LayoutMatch(
                found=True,
                layout_name=best_match[0],
                layout_index=best_match[1],
                confidence=best_score,
            )

        # No match - return alternatives
        return LayoutMatch(
            found=False,
            alternatives=self._layouts,
        )

    def _similarity(self, a: str, b: str) -> float:
        """Calculate similarity between two strings using character n-grams."""
        def ngrams(s: str, n: int = 2) -> set:
            return {s[i:i + n] for i in range(max(1, len(s) - n + 1))}

        a_ngrams = ngrams(a)
        b_ngrams = ngrams(b)

        if not a_ngrams and not b_ngrams:
            return 1.0
        if not a_ngrams or not b_ngrams:
            return 0.0

        intersection = len(a_ngrams & b_ngrams)
        union = len(a_ngrams | b_ngrams)

        return intersection / union


def find_best_layout_match(
    query: str,
    layout_names: List[str],
) -> Tuple[Optional[str], Optional[int], float]:
    """Find the best matching layout name.

    Args:
        query: User-provided layout name
        layout_names: List of available layout names

    Returns:
        Tuple of (layout_name, index, confidence)
        Returns (None, None, 0.0) if no match found.

    Example:
        >>> name, idx, conf = find_best_layout_match("title", layouts)
        >>> if name:
        ...     print(f"Matched: {name} (index {idx}, {conf:.0%} confidence)")
    """
    matcher = LayoutMatcher(layout_names)
    result = matcher.match(query)

    if result.found:
        return (result.layout_name, result.layout_index, result.confidence)
    return (None, None, 0.0)
