"""Template analysis engine.

Provides deep introspection of PowerPoint templates to understand
their structure, layouts, and style elements.
"""

from __future__ import annotations

import json
from dataclasses import dataclass, field, asdict
from pathlib import Path
from typing import Any, Dict, List, Optional, Union

from ..core.presentation import Presentation
from ..oxml.layout import get_layout_info_list, LayoutInfo
from ..oxml.theme import get_theme_part


@dataclass
class PlaceholderAnalysis:
    """Analysis of a single placeholder."""

    type: str
    idx: Optional[int]
    name: str
    bounds: Dict[str, int]  # x, y, cx, cy in EMUs


@dataclass
class LayoutAnalysis:
    """Analysis of a single layout."""

    name: str
    index: int
    placeholders: Dict[str, PlaceholderAnalysis] = field(default_factory=dict)


@dataclass
class ThemeAnalysis:
    """Analysis of theme colors and fonts."""

    colors: Dict[str, str] = field(default_factory=dict)  # name -> hex
    fonts: Dict[str, str] = field(default_factory=dict)  # role -> font name


@dataclass
class TemplateAnalysis:
    """Complete analysis of a PowerPoint template."""

    layouts: Dict[str, LayoutAnalysis] = field(default_factory=dict)
    theme: ThemeAnalysis = field(default_factory=ThemeAnalysis)
    slide_size: Dict[str, int] = field(default_factory=dict)  # width, height
    custom_layouts: List[str] = field(default_factory=list)

    def to_dict(self) -> Dict[str, Any]:
        """Convert to dictionary for JSON serialization."""
        return asdict(self)

    def to_json(self, indent: int = 2) -> str:
        """Convert to JSON string."""
        return json.dumps(self.to_dict(), indent=indent)

    @classmethod
    def from_dict(cls, data: Dict[str, Any]) -> "TemplateAnalysis":
        """Create from dictionary."""
        analysis = cls()

        # Layouts
        for name, layout_data in data.get("layouts", {}).items():
            placeholders = {}
            for ph_name, ph_data in layout_data.get("placeholders", {}).items():
                placeholders[ph_name] = PlaceholderAnalysis(
                    type=ph_data["type"],
                    idx=ph_data.get("idx"),
                    name=ph_data["name"],
                    bounds=ph_data["bounds"],
                )
            analysis.layouts[name] = LayoutAnalysis(
                name=layout_data["name"],
                index=layout_data["index"],
                placeholders=placeholders,
            )

        # Theme
        theme_data = data.get("theme", {})
        analysis.theme = ThemeAnalysis(
            colors=theme_data.get("colors", {}),
            fonts=theme_data.get("fonts", {}),
        )

        # Slide size
        analysis.slide_size = data.get("slide_size", {})

        # Custom layouts
        analysis.custom_layouts = data.get("custom_layouts", [])

        return analysis


def analyze_template(
    template_path: Union[str, Path],
) -> TemplateAnalysis:
    """Analyze a PowerPoint template in depth.

    This function performs comprehensive analysis of a template file,
    extracting information about:
    - All available layouts and their placeholders
    - Theme colors and fonts
    - Slide dimensions
    - Custom/non-standard layouts

    Args:
        template_path: Path to the .pptx template file

    Returns:
        TemplateAnalysis object with all template information

    Example:
        >>> analysis = analyze_template("corporate.pptx")
        >>> print(analysis.theme.colors["accent1"])
        >>> for name, layout in analysis.layouts.items():
        ...     print(f"{name}: {list(layout.placeholders.keys())}")
    """
    # Open the template
    pres = Presentation.open(template_path)

    analysis = TemplateAnalysis()

    # Analyze layouts
    layout_infos = get_layout_info_list(pres._package)

    standard_layouts = {
        "title slide", "title and content", "section header",
        "two content", "comparison", "title only", "blank",
        "content with caption", "picture with caption",
    }

    for layout_info in layout_infos:
        # Build placeholder analysis
        placeholders = {}
        for ph in layout_info.placeholders:
            ph_key = ph.type
            if ph.idx is not None and ph.idx > 0:
                ph_key = f"{ph.type}_{ph.idx}"

            placeholders[ph_key] = PlaceholderAnalysis(
                type=ph.type,
                idx=ph.idx,
                name=ph.name,
                bounds={
                    "x": ph.position.x,
                    "y": ph.position.y,
                    "cx": ph.position.cx,
                    "cy": ph.position.cy,
                },
            )

        layout_analysis = LayoutAnalysis(
            name=layout_info.name,
            index=layout_info.index,
            placeholders=placeholders,
        )

        analysis.layouts[layout_info.name] = layout_analysis

        # Check if custom layout
        if layout_info.name.lower() not in standard_layouts:
            analysis.custom_layouts.append(layout_info.name)

    # Analyze theme
    theme = get_theme_part(pres._package)
    if theme:
        colors = theme.get_colors()
        analysis.theme.colors = {name: f"#{rgb}" for name, rgb in colors.items()}

        fonts = theme.get_fonts()
        analysis.theme.fonts = {
            "heading": fonts.major_font.typeface,
            "body": fonts.minor_font.typeface,
        }

    # Slide size
    width, height = pres._presentation.get_slide_size()
    analysis.slide_size = {
        "width": width,
        "height": height,
        "width_inches": width / 914400,
        "height_inches": height / 914400,
    }

    return analysis


def export_template_schema(
    analysis: TemplateAnalysis,
    output_path: Union[str, Path],
) -> None:
    """Export template analysis to a JSON file.

    This can be used to:
    - Document corporate templates
    - Include in AI system prompts
    - Share template specifications

    Args:
        analysis: TemplateAnalysis object
        output_path: Path for output JSON file

    Example:
        >>> analysis = analyze_template("corporate.pptx")
        >>> export_template_schema(analysis, "corporate_schema.json")
    """
    output_path = Path(output_path)
    with open(output_path, "w") as f:
        f.write(analysis.to_json())


def load_template_schema(
    schema_path: Union[str, Path],
) -> TemplateAnalysis:
    """Load template analysis from a JSON file.

    Args:
        schema_path: Path to JSON schema file

    Returns:
        TemplateAnalysis object

    Example:
        >>> analysis = load_template_schema("corporate_schema.json")
    """
    schema_path = Path(schema_path)
    with open(schema_path) as f:
        data = json.load(f)
    return TemplateAnalysis.from_dict(data)
