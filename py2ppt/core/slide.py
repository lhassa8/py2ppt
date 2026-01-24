"""High-level Slide class."""

from __future__ import annotations

from typing import TYPE_CHECKING, Any

from ..oxml.shapes import (
    Picture,
    Position,
    Shape,
    Table,
    TableCell,
    TextFrame,
)
from ..oxml.slide import SlidePart, update_slide_in_package
from ..utils.colors import is_theme_color, parse_color
from ..utils.errors import PlaceholderNotFoundError
from ..utils.units import parse_length

if TYPE_CHECKING:
    from .presentation import Presentation


class Slide:
    """High-level slide object.

    Provides methods for manipulating slide content.
    """

    def __init__(
        self,
        slide_part: SlidePart,
        slide_number: int,
        presentation: Presentation,
    ) -> None:
        self._part = slide_part
        self._number = slide_number
        self._presentation = presentation

    @property
    def number(self) -> int:
        """Get the slide number (1-indexed)."""
        return self._number

    @property
    def shapes(self) -> list[Shape | Picture | Table]:
        """Get all shapes on the slide."""
        return self._part.shape_tree.shapes

    def get_placeholder_types(self) -> list[str]:
        """Get list of placeholder types on this slide."""
        types = []
        for shape in self._part.shape_tree.get_placeholders():
            if shape.placeholder:
                ph_type = shape.placeholder.type or "body"
                if ph_type not in types:
                    types.append(ph_type)
        return types

    def get_placeholders(self) -> dict[str, Shape]:
        """Get all placeholders as a dict of type -> shape."""
        placeholders = {}
        for shape in self._part.shape_tree.get_placeholders():
            if shape.placeholder:
                ph_type = shape.placeholder.type or "body"
                # Use idx suffix if multiple of same type
                key = ph_type
                idx = shape.placeholder.idx
                if idx is not None and idx > 0:
                    key = f"{ph_type}_{idx}"
                if key not in placeholders:
                    placeholders[key] = shape
        return placeholders

    def _find_placeholder(self, placeholder: str) -> Shape:
        """Find a placeholder by type/name.

        Args:
            placeholder: Placeholder type ("title", "body", "subtitle")
                        or index-suffixed name ("body_1", "body_2")

        Returns:
            The placeholder Shape

        Raises:
            PlaceholderNotFoundError: If placeholder not found
        """
        # Normalize placeholder name
        ph_lower = placeholder.lower().replace("-", "_")

        # Handle aliases
        aliases = {
            "content": "body",
            "text": "body",
            "bullets": "body",
            "sub_title": "subTitle",
            "subtitle": "subTitle",
            "center_title": "ctrTitle",
            "centered_title": "ctrTitle",
            "ctrtitle": "ctrTitle",
        }
        ph_type = aliases.get(ph_lower, ph_lower)

        # Check for indexed placeholder (e.g., "body_1")
        idx = None
        if "_" in ph_type:
            parts = ph_type.rsplit("_", 1)
            if parts[1].isdigit():
                ph_type = parts[0]
                idx = int(parts[1])

        # Search for placeholder
        shape = self._part.shape_tree.get_placeholder(ph_type=ph_type, ph_idx=idx)

        if shape is None:
            # Try without idx
            shape = self._part.shape_tree.get_placeholder(ph_type=ph_type)

        if shape is None:
            available = self.get_placeholder_types()
            raise PlaceholderNotFoundError(placeholder, self._number, available)

        return shape

    def get_title(self) -> str | None:
        """Get the title text."""
        shape = self._part.get_title_placeholder()
        if shape and shape.text_frame:
            return shape.text_frame.text
        return None

    def set_title(
        self,
        text: str,
        *,
        font_size: int | None = None,
        font_family: str | None = None,
        bold: bool = False,
        color: str | None = None,
    ) -> None:
        """Set the title text.

        Args:
            text: Title text
            font_size: Font size in points
            font_family: Font family name
            bold: Whether to make text bold
            color: Color as hex, rgb, or name
        """
        shape = self._part.get_title_placeholder()
        if shape is None:
            raise PlaceholderNotFoundError("title", self._number, self.get_placeholder_types())

        self._set_text_content(shape, text, font_size, font_family, bold, color)
        self._save()

    def get_subtitle(self) -> str | None:
        """Get the subtitle text."""
        shape = self._part.get_subtitle_placeholder()
        if shape and shape.text_frame:
            return shape.text_frame.text
        return None

    def set_subtitle(
        self,
        text: str,
        *,
        font_size: int | None = None,
        font_family: str | None = None,
        bold: bool = False,
        color: str | None = None,
    ) -> None:
        """Set the subtitle text.

        Args:
            text: Subtitle text
            font_size: Font size in points
            font_family: Font family name
            bold: Whether to make text bold
            color: Color as hex, rgb, or name
        """
        shape = self._part.get_subtitle_placeholder()
        if shape is None:
            raise PlaceholderNotFoundError("subtitle", self._number, self.get_placeholder_types())

        self._set_text_content(shape, text, font_size, font_family, bold, color)
        self._save()

    def get_body(self) -> list[str]:
        """Get body content as list of bullet points."""
        shape = self._part.get_body_placeholder()
        if shape is None or shape.text_frame is None:
            return []

        return [p.text for p in shape.text_frame.paragraphs if p.text]

    def set_body(
        self,
        content: str | list[str],
        *,
        levels: list[int] | None = None,
        font_size: int | None = None,
        font_family: str | None = None,
        color: str | None = None,
    ) -> None:
        """Set body content.

        Args:
            content: Single string or list of bullet points
            levels: Optional list of indent levels (0-8)
            font_size: Font size in points
            font_family: Font family name
            color: Color as hex, rgb, or name
        """
        shape = self._part.get_body_placeholder()
        if shape is None:
            raise PlaceholderNotFoundError("body", self._number, self.get_placeholder_types())

        if isinstance(content, str):
            content = [content]

        if levels is None:
            levels = [0] * len(content)

        self._set_bullet_content(shape, content, levels, font_size, font_family, color)
        self._save()

    def add_bullet(
        self,
        text: str,
        *,
        level: int = 0,
        font_size: int | None = None,
        font_family: str | None = None,
        color: str | None = None,
    ) -> None:
        """Add a bullet point to the body.

        Args:
            text: Bullet text
            level: Indent level (0-8)
            font_size: Font size in points
            font_family: Font family name
            color: Color as hex, rgb, or name
        """
        shape = self._part.get_body_placeholder()
        if shape is None:
            raise PlaceholderNotFoundError("body", self._number, self.get_placeholder_types())

        if shape.text_frame is None:
            shape.text_frame = TextFrame()

        # Parse color
        color_val = None
        if color:
            parsed = parse_color(color)
            if not is_theme_color(parsed):
                color_val = parsed

        shape.text_frame.add_paragraph(
            text,
            level=level,
            font_size=font_size,
            color=color_val,
        )
        self._save()

    def set_placeholder_text(
        self,
        placeholder: str,
        text: str,
        *,
        font_size: int | None = None,
        font_family: str | None = None,
        bold: bool = False,
        color: str | None = None,
    ) -> None:
        """Set text in a specific placeholder.

        Args:
            placeholder: Placeholder type or name
            text: Text content
            font_size: Font size in points
            font_family: Font family name
            bold: Whether to make text bold
            color: Color as hex, rgb, or name
        """
        shape = self._find_placeholder(placeholder)
        self._set_text_content(shape, text, font_size, font_family, bold, color)
        self._save()

    def add_text_box(
        self,
        text: str,
        left: str | int,
        top: str | int,
        width: str | int,
        height: str | int,
        *,
        font_size: int | None = None,
        font_family: str | None = None,
        bold: bool = False,
        color: str | None = None,
    ) -> Shape:
        """Add a text box at a specific position.

        Args:
            text: Text content
            left: Left position (e.g., "1in", "2.5cm", or EMU value)
            top: Top position
            width: Width
            height: Height
            font_size: Font size in points
            font_family: Font family name
            bold: Whether to make text bold
            color: Color as hex, rgb, or name

        Returns:
            The new Shape object
        """
        from ..oxml.shapes import create_text_box

        position = Position(
            x=int(parse_length(left)),
            y=int(parse_length(top)),
            cx=int(parse_length(width)),
            cy=int(parse_length(height)),
        )

        color_val = None
        if color:
            parsed = parse_color(color)
            if not is_theme_color(parsed):
                color_val = parsed

        shape = create_text_box(
            shape_id=self._part.shape_tree._next_id,
            text=text,
            position=position,
            font_size=font_size,
            font_family=font_family,
            bold=bold,
            color=color_val,
        )

        self._part.add_shape(shape)
        self._save()

        return shape

    def add_table(
        self,
        data: list[list[Any]],
        left: str | int | None = None,
        top: str | int | None = None,
        width: str | int | None = None,
        height: str | int | None = None,
        *,
        placeholder: str | None = None,
    ) -> Table:
        """Add a table to the slide.

        Args:
            data: 2D list of cell values
            left: Left position (or use placeholder)
            top: Top position
            width: Table width
            height: Table height
            placeholder: Placeholder to fill (alternative to position)

        Returns:
            The new Table object
        """
        if not data or not data[0]:
            raise ValueError("Table data cannot be empty")

        num_rows = len(data)
        num_cols = len(data[0])

        # Determine position
        if placeholder:
            shape = self._find_placeholder(placeholder)
            position = shape.position
        elif left is not None and top is not None:
            position = Position(
                x=int(parse_length(left)),
                y=int(parse_length(top)),
                cx=int(parse_length(width or "6in")),
                cy=int(parse_length(height or "3in")),
            )
        else:
            # Default position
            position = Position(
                x=457200,
                y=1600200,
                cx=8229600,
                cy=3000000,
            )

        # Calculate column widths and row heights
        col_width = position.cx // num_cols
        row_height = position.cy // num_rows

        col_widths = [col_width] * num_cols
        row_heights = [row_height] * num_rows

        # Create cells
        rows = []
        for row_data in data:
            row = []
            for cell_data in row_data:
                cell = TableCell(text=str(cell_data) if cell_data is not None else "")
                row.append(cell)
            rows.append(row)

        table = Table(
            id=self._part.shape_tree._next_id,
            name=f"Table {self._part.shape_tree._next_id}",
            position=position,
            rows=rows,
            col_widths=col_widths,
            row_heights=row_heights,
        )

        self._part.add_table(table)
        self._save()

        return table

    def describe(self) -> dict[str, Any]:
        """Get a description of the slide content.

        Returns:
            Dict with slide information for AI agents
        """
        placeholders = {}
        for ph_type, shape in self.get_placeholders().items():
            if shape.text_frame:
                text = shape.text_frame.text
                if "\n" in text:
                    placeholders[ph_type] = text.split("\n")
                else:
                    placeholders[ph_type] = text

        shapes = []
        for shape in self.shapes:
            if isinstance(shape, Shape):
                if shape.placeholder is None:  # Non-placeholder shapes
                    shapes.append({
                        "type": "text_box",
                        "name": shape.name,
                        "text": shape.text_frame.text if shape.text_frame else "",
                    })
            elif isinstance(shape, Picture):
                shapes.append({
                    "type": "image",
                    "name": shape.name,
                })
            elif isinstance(shape, Table):
                shapes.append({
                    "type": "table",
                    "name": shape.name,
                    "rows": shape.num_rows,
                    "cols": shape.num_cols,
                })

        return {
            "slide_number": self._number,
            "placeholders": placeholders,
            "shapes": shapes,
        }

    def _set_text_content(
        self,
        shape: Shape,
        text: str,
        font_size: int | None,
        font_family: str | None,
        bold: bool,
        color: str | None,
    ) -> None:
        """Set text content in a shape."""
        if shape.text_frame is None:
            shape.text_frame = TextFrame()

        shape.text_frame.clear()

        # Parse color
        color_val = None
        if color:
            parsed = parse_color(color)
            if not is_theme_color(parsed):
                color_val = parsed

        shape.text_frame.add_paragraph(
            text,
            font_size=font_size,
            bold=bold,
            color=color_val,
        )

    def _set_bullet_content(
        self,
        shape: Shape,
        items: list[str],
        levels: list[int],
        font_size: int | None,
        font_family: str | None,
        color: str | None,
    ) -> None:
        """Set bullet content in a shape."""
        if shape.text_frame is None:
            shape.text_frame = TextFrame()

        shape.text_frame.clear()

        # Parse color
        color_val = None
        if color:
            parsed = parse_color(color)
            if not is_theme_color(parsed):
                color_val = parsed

        for item, level in zip(items, levels, strict=False):
            shape.text_frame.add_paragraph(
                item,
                level=level,
                font_size=font_size,
                color=color_val,
            )

    def _save(self) -> None:
        """Save changes to the package."""
        update_slide_in_package(
            self._presentation._package,
            self._number,
            self._part,
        )
