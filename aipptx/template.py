"""Template class for AI-friendly template analysis.

Provides comprehensive analysis of PowerPoint templates with
AI-readable descriptions of layouts, placeholders, and theme colors.
"""

from __future__ import annotations

from pathlib import Path
from typing import Any

from .layout import LayoutDescription, LayoutRecommendation, analyze_layout, recommend_layout


class Template:
    """AI-friendly wrapper for PowerPoint templates.

    Provides template analysis and creates presentations with
    semantic, high-level APIs designed for AI agents.

    Example:
        >>> template = Template("corporate.pptx")
        >>> layouts = template.describe()
        >>> print(layouts[0]["name"], layouts[0]["description"])
        >>> pres = template.create_presentation()
    """

    def __init__(self, template_path: str | Path) -> None:
        """Load and analyze a template.

        Args:
            template_path: Path to the .pptx template file
        """
        self._path = Path(template_path)
        if not self._path.exists():
            raise FileNotFoundError(f"Template not found: {template_path}")

        # Import py2ppt here to avoid circular imports
        from py2ppt.core.presentation import Presentation as Py2PptPresentation
        from py2ppt.oxml.layout import get_layout_info_list
        from py2ppt.oxml.theme import get_theme_part

        # Load the template
        self._pres = Py2PptPresentation.open(self._path)

        # Analyze layouts
        self._layout_infos = get_layout_info_list(self._pres._package)
        self._layouts: list[LayoutDescription] = []
        self._analyze_layouts()

        # Get theme
        self._theme = get_theme_part(self._pres._package)
        self._colors: dict[str, str] = {}
        self._fonts: dict[str, str] = {}
        self._extract_theme()

    def _analyze_layouts(self) -> None:
        """Analyze all layouts and build descriptions."""
        for layout_info in self._layout_infos:
            # Convert placeholder info to dict format
            placeholders = []
            for ph in layout_info.placeholders:
                placeholders.append({
                    "type": ph.type,
                    "idx": ph.idx,
                    "name": ph.name,
                    "x": ph.position.x,
                    "y": ph.position.y,
                    "cx": ph.position.cx,
                    "cy": ph.position.cy,
                })

            layout_desc = analyze_layout(
                name=layout_info.name,
                index=layout_info.index,
                placeholders=placeholders,
            )
            self._layouts.append(layout_desc)

    def _extract_theme(self) -> None:
        """Extract theme colors and fonts."""
        if self._theme:
            colors = self._theme.get_colors()
            self._colors = {name: f"#{rgb}" for name, rgb in colors.items()}

            fonts = self._theme.get_fonts()
            self._fonts = {
                "heading": fonts.major_font.typeface,
                "body": fonts.minor_font.typeface,
            }
        else:
            # Defaults
            self._fonts = {"heading": "Calibri Light", "body": "Calibri"}

    @property
    def path(self) -> Path:
        """Get the template file path."""
        return self._path

    @property
    def colors(self) -> dict[str, str]:
        """Get theme colors.

        Returns:
            Dict mapping color names to hex values.
            Common keys: dk1, lt1, dk2, lt2, accent1-6, hlink, folHlink

        Example:
            >>> colors = template.colors
            >>> primary = colors.get("accent1", "#000000")
        """
        return self._colors.copy()

    @property
    def fonts(self) -> dict[str, str]:
        """Get theme fonts.

        Returns:
            Dict with 'heading' and 'body' font names.

        Example:
            >>> fonts = template.fonts
            >>> print(f"Use {fonts['heading']} for titles")
        """
        return self._fonts.copy()

    def describe(self) -> list[dict[str, Any]]:
        """Get AI-friendly descriptions of all layouts.

        Returns a list of layout descriptions that can be easily
        consumed by AI agents to understand what layouts are available
        and how to use them.

        Returns:
            List of layout description dicts with:
            - name: Layout name (e.g., "Title Slide")
            - index: Layout index for selection
            - type: Classified type (title_slide, content, etc.)
            - description: Human-readable description
            - placeholders: Dict of semantic placeholder names
            - best_for: List of content types this layout suits

        Example:
            >>> layouts = template.describe()
            >>> for layout in layouts:
            ...     print(f"{layout['name']}: {layout['description']}")
            ...     print(f"  Best for: {', '.join(layout['best_for'])}")
        """
        return [layout.to_dict() for layout in self._layouts]

    def describe_as_text(self) -> str:
        """Get a text description of the template for AI prompts.

        Returns a formatted string that can be included in AI system
        prompts to give context about the template.

        Returns:
            Multi-line string describing the template.

        Example:
            >>> description = template.describe_as_text()
            >>> # Include in AI prompt:
            >>> prompt = f"Template info:\\n{description}\\n\\nCreate a presentation..."
        """
        lines = [
            f"Template: {self._path.name}",
            f"Theme Colors: {', '.join(f'{k}={v}' for k, v in list(self._colors.items())[:6])}",
            f"Fonts: heading={self._fonts.get('heading')}, body={self._fonts.get('body')}",
            "",
            "Available Layouts:",
        ]

        for layout in self._layouts:
            ph_names = list(layout.placeholders.keys())
            lines.append(f"  {layout.index}: {layout.name} ({layout.layout_type.value})")
            lines.append(f"      Placeholders: {', '.join(ph_names)}")
            lines.append(f"      Best for: {', '.join(layout.best_for)}")

        return "\n".join(lines)

    def get_layout(self, name_or_index: str | int) -> LayoutDescription | None:
        """Get a specific layout by name or index.

        Args:
            name_or_index: Layout name (fuzzy matched) or index

        Returns:
            LayoutDescription or None if not found

        Example:
            >>> layout = template.get_layout("title")
            >>> if layout:
            ...     print(layout.placeholders)
        """
        if isinstance(name_or_index, int):
            for layout in self._layouts:
                if layout.index == name_or_index:
                    return layout
            return None

        # Fuzzy name match
        name_lower = name_or_index.lower()
        for layout in self._layouts:
            if name_lower in layout.name.lower() or layout.name.lower() in name_lower:
                return layout
        return None

    def recommend_layout(
        self,
        content_type: str,
        has_image: bool = False,
        bullet_count: int = 0,
    ) -> list[dict[str, Any]]:
        """Get layout recommendations for content type.

        Args:
            content_type: Type of content ("bullets", "comparison", etc.)
            has_image: Whether content includes an image
            bullet_count: Number of bullet points

        Returns:
            List of recommendations sorted by confidence

        Example:
            >>> recs = template.recommend_layout("comparison")
            >>> best = recs[0]
            >>> print(f"Use '{best['name']}' ({best['confidence']:.0%} confidence)")
        """
        recommendations = recommend_layout(
            layouts=self._layouts,
            content_type=content_type,
            has_image=has_image,
            bullet_count=bullet_count,
        )

        return [
            {
                "name": r.layout_name,
                "index": r.layout_index,
                "confidence": r.confidence,
                "reason": r.reason,
            }
            for r in recommendations
        ]

    def create_presentation(self) -> "Presentation":
        """Create a new presentation from this template.

        Returns a Presentation object with high-level methods for
        adding slides with semantic APIs.

        Returns:
            Presentation object

        Example:
            >>> pres = template.create_presentation()
            >>> pres.add_title_slide("My Title", "Subtitle")
            >>> pres.save("output.pptx")
        """
        from .presentation import Presentation

        return Presentation(self)

    def get_layout_names(self) -> list[str]:
        """Get list of all layout names.

        Returns:
            List of layout names in order
        """
        return [layout.name for layout in self._layouts]

    def __repr__(self) -> str:
        return f"Template({self._path.name!r}, {len(self._layouts)} layouts)"
