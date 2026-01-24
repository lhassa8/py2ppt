"""Placeholder detection and mapping utilities."""

from __future__ import annotations

from dataclasses import dataclass
from typing import Dict, List, Optional, Tuple


# Placeholder type aliases - maps common names to PowerPoint placeholder types
PLACEHOLDER_ALIASES: Dict[str, List[str]] = {
    # Title variations
    "title": ["title", "ctrTitle"],
    "centered_title": ["ctrTitle"],
    "center_title": ["ctrTitle"],

    # Subtitle variations
    "subtitle": ["subTitle"],
    "sub_title": ["subTitle"],

    # Body/content variations
    "body": ["body", "obj"],
    "content": ["body", "obj"],
    "text": ["body", "obj"],
    "bullets": ["body"],

    # Media
    "picture": ["pic"],
    "image": ["pic"],
    "chart": ["chart"],
    "table": ["tbl"],
    "diagram": ["dgm"],
    "media": ["media"],

    # Metadata
    "date": ["dt"],
    "footer": ["ftr"],
    "slide_number": ["sldNum"],
    "header": ["hdr"],
}

# Reverse mapping: PowerPoint type -> friendly name
PLACEHOLDER_FRIENDLY_NAMES: Dict[str, str] = {
    "title": "title",
    "ctrTitle": "centered_title",
    "subTitle": "subtitle",
    "body": "body",
    "obj": "content",
    "pic": "picture",
    "chart": "chart",
    "tbl": "table",
    "dgm": "diagram",
    "media": "media",
    "dt": "date",
    "ftr": "footer",
    "sldNum": "slide_number",
    "hdr": "header",
    "clipArt": "clip_art",
    "sldImg": "slide_image",
}


def normalize_placeholder_name(name: str) -> str:
    """Normalize a placeholder name for matching.

    Converts user-friendly names to PowerPoint types.

    Examples:
        "title" -> "title"
        "content" -> "body"
        "sub_title" -> "subTitle"
    """
    name = name.lower().strip().replace("-", "_").replace(" ", "_")

    if name in PLACEHOLDER_ALIASES:
        # Return the first (primary) type
        return PLACEHOLDER_ALIASES[name][0]

    return name


def get_placeholder_candidates(name: str) -> List[str]:
    """Get list of PowerPoint placeholder types for a given name.

    Returns multiple candidates for fuzzy matching.

    Examples:
        "title" -> ["title", "ctrTitle"]
        "body" -> ["body", "obj"]
    """
    name = name.lower().strip().replace("-", "_").replace(" ", "_")

    if name in PLACEHOLDER_ALIASES:
        return PLACEHOLDER_ALIASES[name]

    # Not found in aliases, return as-is
    return [name]


def get_friendly_name(ph_type: str) -> str:
    """Get user-friendly name for a PowerPoint placeholder type.

    Examples:
        "ctrTitle" -> "centered_title"
        "body" -> "body"
    """
    return PLACEHOLDER_FRIENDLY_NAMES.get(ph_type, ph_type)


@dataclass
class PlaceholderMatch:
    """Result of a placeholder match operation."""

    found: bool
    placeholder_type: Optional[str] = None
    placeholder_idx: Optional[int] = None
    friendly_name: Optional[str] = None


class PlaceholderMap:
    """Maps user-provided placeholder names to actual placeholders.

    Supports:
    - Exact matching: "title", "body"
    - Alias matching: "content" -> "body"
    - Indexed matching: "body_1", "body_2"
    - Fuzzy matching: case-insensitive, underscore/hyphen variations
    """

    def __init__(self) -> None:
        self._by_type: Dict[str, List[Tuple[str, int | None]]] = {}

    def add(self, ph_type: str, idx: Optional[int] = None) -> None:
        """Register a placeholder.

        Args:
            ph_type: PowerPoint placeholder type
            idx: Placeholder index (for multiple of same type)
        """
        if ph_type not in self._by_type:
            self._by_type[ph_type] = []
        self._by_type[ph_type].append((ph_type, idx))

    def find(self, name: str) -> PlaceholderMatch:
        """Find a placeholder by user-provided name.

        Args:
            name: User-provided placeholder name

        Returns:
            PlaceholderMatch with results
        """
        # Parse indexed names like "body_1"
        idx = None
        base_name = name.lower().strip().replace("-", "_")

        if "_" in base_name:
            parts = base_name.rsplit("_", 1)
            if parts[1].isdigit():
                base_name = parts[0]
                idx = int(parts[1])

        # Get candidate types
        candidates = get_placeholder_candidates(base_name)

        # Search for matches
        for ph_type in candidates:
            if ph_type in self._by_type:
                matches = self._by_type[ph_type]

                # If index specified, find exact match
                if idx is not None:
                    for _, match_idx in matches:
                        if match_idx == idx:
                            return PlaceholderMatch(
                                found=True,
                                placeholder_type=ph_type,
                                placeholder_idx=idx,
                                friendly_name=get_friendly_name(ph_type),
                            )
                else:
                    # Return first match
                    if matches:
                        _, first_idx = matches[0]
                        return PlaceholderMatch(
                            found=True,
                            placeholder_type=ph_type,
                            placeholder_idx=first_idx,
                            friendly_name=get_friendly_name(ph_type),
                        )

        return PlaceholderMatch(found=False)

    def list_all(self) -> List[str]:
        """List all registered placeholders with friendly names."""
        result = []
        for ph_type, entries in self._by_type.items():
            friendly = get_friendly_name(ph_type)
            if len(entries) == 1:
                result.append(friendly)
            else:
                for i, (_, idx) in enumerate(entries):
                    if idx is not None:
                        result.append(f"{friendly}_{idx}")
                    else:
                        result.append(f"{friendly}_{i}")
        return result

    @classmethod
    def from_shape_tree(cls, shape_tree) -> "PlaceholderMap":
        """Create a PlaceholderMap from a shape tree.

        Args:
            shape_tree: ShapeTree object

        Returns:
            Populated PlaceholderMap
        """
        pm = cls()
        for shape in shape_tree.get_placeholders():
            if shape.placeholder:
                pm.add(
                    shape.placeholder.type or "body",
                    shape.placeholder.idx,
                )
        return pm
