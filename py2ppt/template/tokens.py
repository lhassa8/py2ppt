"""Design token system for brand consistency.

Design tokens are named values for colors, fonts, and other design
properties that ensure consistent branding across presentations.
"""

from __future__ import annotations

import json
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any, Dict, Optional, Union

from ..utils.colors import parse_color, is_valid_hex_color


@dataclass
class FontToken:
    """Font token specification."""

    family: str
    size: Optional[str] = None  # e.g., "32pt"
    weight: Optional[str] = None  # "normal", "bold"


@dataclass
class DesignTokens:
    """Collection of design tokens for a brand/template.

    Design tokens provide named values that can be used throughout
    presentations to ensure brand consistency.
    """

    colors: Dict[str, str] = field(default_factory=dict)  # name -> hex
    fonts: Dict[str, FontToken] = field(default_factory=dict)  # name -> FontToken
    spacing: Dict[str, str] = field(default_factory=dict)  # name -> measurement
    metadata: Dict[str, Any] = field(default_factory=dict)

    def get_color(self, name: str) -> Optional[str]:
        """Get a color by token name.

        Args:
            name: Token name (e.g., "brand-primary")

        Returns:
            Hex color string or None if not found
        """
        return self.colors.get(name)

    def get_font(self, name: str) -> Optional[FontToken]:
        """Get a font by token name.

        Args:
            name: Token name (e.g., "heading")

        Returns:
            FontToken or None if not found
        """
        return self.fonts.get(name)

    def resolve_color(self, color: str) -> str:
        """Resolve a color value, looking up tokens if needed.

        Args:
            color: Color value or token name

        Returns:
            Hex color string

        Raises:
            ValueError: If color cannot be resolved
        """
        # Check if it's a token name
        if color in self.colors:
            return self.colors[color]

        # Try to parse as a direct color value
        return "#" + parse_color(color)

    def to_dict(self) -> Dict[str, Any]:
        """Convert to dictionary for JSON serialization."""
        return {
            "colors": self.colors,
            "fonts": {
                name: {
                    "family": font.family,
                    "size": font.size,
                    "weight": font.weight,
                }
                for name, font in self.fonts.items()
            },
            "spacing": self.spacing,
            "metadata": self.metadata,
        }

    @classmethod
    def from_dict(cls, data: Dict[str, Any]) -> "DesignTokens":
        """Create from dictionary."""
        tokens = cls()

        # Colors
        tokens.colors = data.get("colors", {})

        # Fonts
        for name, font_data in data.get("fonts", {}).items():
            if isinstance(font_data, str):
                tokens.fonts[name] = FontToken(family=font_data)
            else:
                tokens.fonts[name] = FontToken(
                    family=font_data.get("family", "Arial"),
                    size=font_data.get("size"),
                    weight=font_data.get("weight"),
                )

        # Spacing
        tokens.spacing = data.get("spacing", {})

        # Metadata
        tokens.metadata = data.get("metadata", {})

        return tokens


def create_tokens(spec: Dict[str, Any]) -> DesignTokens:
    """Create design tokens from a specification dict.

    Args:
        spec: Token specification with colors, fonts, spacing

    Returns:
        DesignTokens object

    Example:
        >>> tokens = create_tokens({
        ...     "colors": {
        ...         "brand-primary": "#0066CC",
        ...         "brand-secondary": "#FF6600",
        ...         "text-dark": "#333333",
        ...     },
        ...     "fonts": {
        ...         "heading": {"family": "Arial Black", "size": "32pt"},
        ...         "body": {"family": "Arial", "size": "14pt"},
        ...     }
        ... })
    """
    tokens = DesignTokens.from_dict(spec)

    # Validate colors
    for name, color in tokens.colors.items():
        if not color.startswith("#"):
            tokens.colors[name] = "#" + parse_color(color)

    return tokens


def save_tokens(
    tokens: DesignTokens,
    path: Union[str, Path],
) -> None:
    """Save design tokens to a JSON file.

    Args:
        tokens: DesignTokens to save
        path: Output file path

    Example:
        >>> save_tokens(tokens, "brand_tokens.json")
    """
    path = Path(path)
    with open(path, "w") as f:
        json.dump(tokens.to_dict(), f, indent=2)


def load_tokens(path: Union[str, Path]) -> DesignTokens:
    """Load design tokens from a JSON file.

    Args:
        path: Path to tokens JSON file

    Returns:
        DesignTokens object

    Example:
        >>> tokens = load_tokens("brand_tokens.json")
    """
    path = Path(path)
    with open(path) as f:
        data = json.load(f)
    return DesignTokens.from_dict(data)


# Predefined token sets

DEFAULT_TOKENS = DesignTokens(
    colors={
        "primary": "#4472C4",
        "secondary": "#ED7D31",
        "success": "#70AD47",
        "warning": "#FFC000",
        "error": "#C00000",
        "text-dark": "#333333",
        "text-light": "#666666",
        "background": "#FFFFFF",
    },
    fonts={
        "heading": FontToken(family="Calibri Light", size="44pt"),
        "subheading": FontToken(family="Calibri Light", size="32pt"),
        "body": FontToken(family="Calibri", size="18pt"),
        "caption": FontToken(family="Calibri", size="12pt"),
    },
)


def get_default_tokens() -> DesignTokens:
    """Get the default token set."""
    return DEFAULT_TOKENS
