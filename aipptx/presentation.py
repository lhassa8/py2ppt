"""Presentation class with semantic slide management.

Provides high-level methods for creating presentations with
AI-friendly, intent-based APIs.
"""

from __future__ import annotations

from pathlib import Path
from typing import TYPE_CHECKING, Any

from .formatting import format_for_py2ppt, parse_content
from .layout import LayoutType

if TYPE_CHECKING:
    from .template import Template


class Presentation:
    """AI-friendly presentation with semantic slide methods.

    This class wraps py2ppt's Presentation to provide high-level,
    intent-based methods for creating slides. Instead of dealing
    with placeholder indices and layout details, you can simply
    call methods like add_title_slide() or add_comparison_slide().

    Example:
        >>> template = Template("corporate.pptx")
        >>> pres = template.create_presentation()
        >>>
        >>> pres.add_title_slide("Q4 Review", "January 2025")
        >>> pres.add_content_slide("Key Points", [
        ...     "Revenue up 20%",
        ...     "New markets opened",
        ...     "Customer satisfaction high"
        ... ])
        >>> pres.save("output.pptx")
    """

    def __init__(self, template: "Template") -> None:
        """Initialize from a Template.

        Use Template.create_presentation() instead of this constructor.
        """
        self._template = template
        self._layouts = template._layouts

        # Create py2ppt presentation from template
        from py2ppt.core.presentation import Presentation as Py2PptPresentation

        self._pres = Py2PptPresentation.from_template(template.path)

    @property
    def slide_count(self) -> int:
        """Get the current number of slides."""
        return self._pres.slide_count

    @property
    def template(self) -> "Template":
        """Get the template this presentation was created from."""
        return self._template

    def _find_layout_by_type(self, layout_type: LayoutType) -> int:
        """Find the best layout index for a layout type."""
        for layout in self._layouts:
            if layout.layout_type == layout_type:
                return layout.index
        # Fall back to first layout
        return 0

    def _find_layout(self, layout: str | int | None, layout_type: LayoutType) -> int:
        """Find layout by name, index, or type."""
        if layout is None or layout == "auto":
            return self._find_layout_by_type(layout_type)

        if isinstance(layout, int):
            return layout

        # Fuzzy name match
        layout_lower = layout.lower()
        for l in self._layouts:
            if layout_lower in l.name.lower() or l.name.lower() in layout_lower:
                return l.index
        return self._find_layout_by_type(layout_type)

    def add_title_slide(
        self,
        title: str,
        subtitle: str = "",
        *,
        layout: str | int | None = None,
    ) -> int:
        """Add a title/cover slide.

        Args:
            title: Main presentation title
            subtitle: Optional subtitle (presenter name, date, etc.)
            layout: Layout name, index, or None for auto-selection

        Returns:
            Slide number of the new slide

        Example:
            >>> pres.add_title_slide("Q4 Business Review", "January 2025")
        """
        layout_idx = self._find_layout(layout, LayoutType.TITLE)
        slide = self._pres.add_slide(layout=layout_idx)
        slide_num = slide.number

        # Set title
        import py2ppt

        py2ppt.set_title(self._pres, slide_num, title)

        # Set subtitle if provided
        if subtitle:
            py2ppt.set_subtitle(self._pres, slide_num, subtitle)

        return slide_num

    def add_section_slide(
        self,
        title: str,
        subtitle: str = "",
        *,
        layout: str | int | None = None,
    ) -> int:
        """Add a section header/divider slide.

        Args:
            title: Section title
            subtitle: Optional subtitle or description
            layout: Layout name, index, or None for auto-selection

        Returns:
            Slide number of the new slide

        Example:
            >>> pres.add_section_slide("Part 2: Analysis")
        """
        layout_idx = self._find_layout(layout, LayoutType.SECTION)
        slide = self._pres.add_slide(layout=layout_idx)
        slide_num = slide.number

        import py2ppt

        py2ppt.set_title(self._pres, slide_num, title)

        if subtitle:
            try:
                py2ppt.set_subtitle(self._pres, slide_num, subtitle)
            except Exception:
                # Section layouts may not have subtitle placeholder
                pass

        return slide_num

    def add_content_slide(
        self,
        title: str,
        content: str | list[str | tuple | dict | list],
        *,
        levels: list[int] | None = None,
        layout: str | int | None = None,
    ) -> int:
        """Add a content slide with bullets.

        Args:
            title: Slide title
            content: Content as string, list of strings, or rich text
            levels: Optional indent levels for each bullet (0=top level)
            layout: Layout name, index, or None for auto-selection

        Returns:
            Slide number of the new slide

        Example:
            >>> # Simple bullets
            >>> pres.add_content_slide("Key Points", [
            ...     "First point",
            ...     "Second point",
            ...     "Third point"
            ... ])

            >>> # With nested levels
            >>> pres.add_content_slide("Details", [
            ...     "Main topic",
            ...     "Sub-point 1",
            ...     "Sub-point 2",
            ...     "Another main topic"
            ... ], levels=[0, 1, 1, 0])

            >>> # Rich text
            >>> pres.add_content_slide("Important", [
            ...     [{"text": "Key: ", "bold": True}, {"text": "value"}],
            ...     "Normal bullet"
            ... ])
        """
        layout_idx = self._find_layout(layout, LayoutType.CONTENT)
        slide = self._pres.add_slide(layout=layout_idx)
        slide_num = slide.number

        import py2ppt

        py2ppt.set_title(self._pres, slide_num, title)

        # Parse and format content
        paragraphs = parse_content(content, levels)
        formatted_content, formatted_levels = format_for_py2ppt(paragraphs)

        py2ppt.set_body(self._pres, slide_num, formatted_content, levels=formatted_levels)

        return slide_num

    def add_two_column_slide(
        self,
        title: str,
        left_content: list[str | tuple | dict | list],
        right_content: list[str | tuple | dict | list],
        *,
        left_levels: list[int] | None = None,
        right_levels: list[int] | None = None,
        layout: str | int | None = None,
    ) -> int:
        """Add a two-column content slide.

        Args:
            title: Slide title
            left_content: Content for left column
            right_content: Content for right column
            left_levels: Indent levels for left content
            right_levels: Indent levels for right content
            layout: Layout name, index, or None for auto-selection

        Returns:
            Slide number of the new slide

        Example:
            >>> pres.add_two_column_slide(
            ...     "Topics",
            ...     ["Point A", "Point B"],
            ...     ["Point X", "Point Y"]
            ... )
        """
        layout_idx = self._find_layout(layout, LayoutType.TWO_COLUMN)
        slide = self._pres.add_slide(layout=layout_idx)
        slide_num = slide.number

        import py2ppt

        py2ppt.set_title(self._pres, slide_num, title)

        # Format content
        left_paragraphs = parse_content(left_content, left_levels)
        left_formatted, left_lvls = format_for_py2ppt(left_paragraphs)

        right_paragraphs = parse_content(right_content, right_levels)
        right_formatted, right_lvls = format_for_py2ppt(right_paragraphs)

        # Set left content (body_1 or first body placeholder)
        try:
            py2ppt.set_placeholder_text(
                self._pres, slide_num, "body_1", "\n".join(str(c) for c in left_formatted)
            )
        except Exception:
            # Try setting via set_body which targets first body
            py2ppt.set_body(self._pres, slide_num, left_formatted, levels=left_lvls)

        # Set right content (body_2)
        try:
            py2ppt.set_placeholder_text(
                self._pres, slide_num, "body_2", "\n".join(str(c) for c in right_formatted)
            )
        except Exception:
            pass  # Layout may not support second body

        return slide_num

    def add_comparison_slide(
        self,
        title: str,
        left_heading: str,
        left_content: list[str | tuple | dict | list],
        right_heading: str,
        right_content: list[str | tuple | dict | list],
        *,
        layout: str | int | None = None,
    ) -> int:
        """Add a comparison slide with two titled columns.

        Ideal for before/after, pros/cons, or versus comparisons.

        Args:
            title: Slide title
            left_heading: Heading for left column
            left_content: Bullet points for left column
            right_heading: Heading for right column
            right_content: Bullet points for right column
            layout: Layout name, index, or None for auto-selection

        Returns:
            Slide number of the new slide

        Example:
            >>> pres.add_comparison_slide(
            ...     "Before vs After",
            ...     "Legacy System",
            ...     ["Slow", "Manual", "Error-prone"],
            ...     "New Platform",
            ...     ["Fast", "Automated", "Reliable"]
            ... )
        """
        # Try to find a comparison layout, fall back to two-column
        layout_idx = self._find_layout(layout, LayoutType.COMPARISON)
        slide = self._pres.add_slide(layout=layout_idx)
        slide_num = slide.number

        import py2ppt

        py2ppt.set_title(self._pres, slide_num, title)

        # Get placeholders to understand the layout
        # get_placeholders returns dict[str, str] mapping type to content
        placeholders = py2ppt.get_placeholders(self._pres, slide_num)

        # Count body placeholders
        body_ph_count = sum(1 for k in placeholders.keys() if "body" in k.lower())

        if body_ph_count >= 4:
            # True comparison layout with 4 body placeholders
            try:
                py2ppt.set_placeholder_text(self._pres, slide_num, "body_1", left_heading)
                left_paragraphs = parse_content(left_content)
                left_formatted, _ = format_for_py2ppt(left_paragraphs)
                left_text = "\n".join(
                    c if isinstance(c, str) else (c[0]["text"] if isinstance(c, list) else str(c))
                    for c in left_formatted
                )
                py2ppt.set_placeholder_text(self._pres, slide_num, "body_2", left_text)

                py2ppt.set_placeholder_text(self._pres, slide_num, "body_3", right_heading)
                right_paragraphs = parse_content(right_content)
                right_formatted, _ = format_for_py2ppt(right_paragraphs)
                right_text = "\n".join(
                    c if isinstance(c, str) else (c[0]["text"] if isinstance(c, list) else str(c))
                    for c in right_formatted
                )
                py2ppt.set_placeholder_text(self._pres, slide_num, "body_4", right_text)
            except Exception:
                pass
        elif body_ph_count >= 2:
            # Two-column layout - combine heading with content
            left_combined = [{"text": left_heading, "bold": True}] + list(left_content)
            right_combined = [{"text": right_heading, "bold": True}] + list(right_content)

            left_levels = [0] + [1] * len(left_content)
            right_levels = [0] + [1] * len(right_content)

            left_paragraphs = parse_content(left_combined, left_levels)
            right_paragraphs = parse_content(right_combined, right_levels)

            left_formatted, left_lvls = format_for_py2ppt(left_paragraphs)
            right_formatted, right_lvls = format_for_py2ppt(right_paragraphs)

            try:
                py2ppt.set_placeholder_text(
                    self._pres, slide_num, "body_1",
                    "\n".join(c if isinstance(c, str) else str(c) for c in left_formatted)
                )
                py2ppt.set_placeholder_text(
                    self._pres, slide_num, "body_2",
                    "\n".join(c if isinstance(c, str) else str(c) for c in right_formatted)
                )
            except Exception:
                pass
        else:
            # Single body placeholder - combine all content
            combined = [
                f"**{left_heading}**",
                *[f"  - {c}" if isinstance(c, str) else str(c) for c in left_content],
                "",
                f"**{right_heading}**",
                *[f"  - {c}" if isinstance(c, str) else str(c) for c in right_content],
            ]
            try:
                py2ppt.set_body(self._pres, slide_num, combined)
            except Exception:
                pass

        return slide_num

    def add_image_slide(
        self,
        title: str,
        image_path: str | Path,
        caption: str = "",
        *,
        layout: str | int | None = None,
    ) -> int:
        """Add a slide with an image.

        Args:
            title: Slide title
            image_path: Path to the image file
            caption: Optional caption text
            layout: Layout name, index, or None for auto-selection

        Returns:
            Slide number of the new slide

        Example:
            >>> pres.add_image_slide(
            ...     "Product Photo",
            ...     "product.png",
            ...     "Our flagship product"
            ... )
        """
        layout_idx = self._find_layout(layout, LayoutType.IMAGE_CONTENT)
        slide = self._pres.add_slide(layout=layout_idx)
        slide_num = slide.number

        import py2ppt

        py2ppt.set_title(self._pres, slide_num, title)

        # Add image - use sensible defaults for positioning
        py2ppt.add_image(
            self._pres,
            slide_num,
            str(image_path),
            left="1in",
            top="2in",
            width="5in",
        )

        if caption:
            # Try to set in body placeholder
            try:
                py2ppt.set_body(self._pres, slide_num, [caption])
            except Exception:
                pass

        return slide_num

    def add_blank_slide(self, layout: str | int | None = None) -> int:
        """Add a blank slide.

        Args:
            layout: Layout name, index, or None for auto-selection

        Returns:
            Slide number of the new slide

        Example:
            >>> slide_num = pres.add_blank_slide()
        """
        layout_idx = self._find_layout(layout, LayoutType.BLANK)
        slide = self._pres.add_slide(layout=layout_idx)
        return slide.number

    def add_slide(
        self,
        layout: str | int = "auto",
        content_type: str = "content",
        title: str = "",
        content: list[str] | None = None,
        **kwargs: Any,
    ) -> int:
        """Add a slide with auto-layout selection.

        A flexible method that chooses the right slide type
        based on content_type and delegates to the appropriate
        specialized method.

        Args:
            layout: Layout name, index, or "auto"
            content_type: Type of content ("title", "content", "comparison", etc.)
            title: Slide title
            content: Content items
            **kwargs: Additional arguments for the specific slide type

        Returns:
            Slide number of the new slide

        Example:
            >>> pres.add_slide(
            ...     content_type="bullets",
            ...     title="Key Points",
            ...     content=["Point 1", "Point 2"]
            ... )
        """
        content_lower = content_type.lower()

        if content_lower in ("title", "cover", "opening"):
            subtitle = kwargs.get("subtitle", "")
            return self.add_title_slide(title, subtitle, layout=layout)

        elif content_lower in ("section", "divider"):
            subtitle = kwargs.get("subtitle", "")
            return self.add_section_slide(title, subtitle, layout=layout)

        elif content_lower in ("comparison", "versus", "vs"):
            return self.add_comparison_slide(
                title,
                kwargs.get("left_heading", "Option A"),
                kwargs.get("left_content", content or []),
                kwargs.get("right_heading", "Option B"),
                kwargs.get("right_content", []),
                layout=layout,
            )

        elif content_lower in ("two_column", "split"):
            return self.add_two_column_slide(
                title,
                kwargs.get("left_content", content or []),
                kwargs.get("right_content", []),
                layout=layout,
            )

        elif content_lower in ("image", "picture"):
            return self.add_image_slide(
                title,
                kwargs.get("image_path", ""),
                kwargs.get("caption", ""),
                layout=layout,
            )

        elif content_lower == "blank":
            return self.add_blank_slide(layout=layout)

        else:
            # Default to content slide
            return self.add_content_slide(
                title,
                content or [],
                levels=kwargs.get("levels"),
                layout=layout,
            )

    def set_notes(self, slide_number: int, notes: str) -> None:
        """Set speaker notes for a slide.

        Args:
            slide_number: The slide number (1-indexed)
            notes: Notes text (can include newlines)

        Example:
            >>> pres.set_notes(1, "Remember to mention...")
        """
        import py2ppt

        py2ppt.set_notes(self._pres, slide_number, notes)

    def save(self, path: str | Path) -> None:
        """Save the presentation.

        Args:
            path: Output file path

        Example:
            >>> pres.save("output.pptx")
        """
        self._pres.save(path)

    def __repr__(self) -> str:
        return f"Presentation({self.slide_count} slides, template={self._template.path.name})"
