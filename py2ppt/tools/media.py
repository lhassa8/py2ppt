"""Media tool functions.

Functions for adding tables, images, charts, and other media to slides.
"""

from __future__ import annotations

from pathlib import Path
from typing import Any, Literal

from ..core.presentation import Presentation
from ..oxml.chart import ChartData, ChartPart, ChartProperties, ChartSeries
from ..oxml.ns import CONTENT_TYPE, REL_TYPE
from ..oxml.shapes import (
    BorderStyle,
    CellStyle,
    Chart,
    CropRect,
    Picture,
    PictureEffects,
    Position,
    Table,
)
from ..utils.units import parse_length


def add_table(
    presentation: Presentation,
    slide_number: int,
    data: list[list[Any]],
    *,
    placeholder: str | None = None,
    left: str | int | None = None,
    top: str | int | None = None,
    width: str | int | None = None,
    height: str | int | None = None,
    header_row: bool = True,
    banded_rows: bool = True,
    first_col: bool = False,
    last_col: bool = False,
    header_background: str | None = None,
    header_text_color: str | None = None,
    cell_borders: bool = True,
    border_color: str = "#000000",
) -> None:
    """Add a table to a slide.

    Tables use theme styling by default for headers and banding.
    Override with explicit colors for custom branding.

    Args:
        presentation: The presentation to modify
        slide_number: The slide number (1-indexed)
        data: 2D list of cell values. First row is typically headers.
        placeholder: Placeholder to fill (e.g., "content", "body").
                    If provided, table will use placeholder's position.
        left: Left position (e.g., "1in", "2.5cm") - ignored if placeholder set
        top: Top position - ignored if placeholder set
        width: Table width - ignored if placeholder set
        height: Table height - ignored if placeholder set
        header_row: Style first row as header (default True)
        banded_rows: Alternate row colors (default True)
        first_col: Style first column differently
        last_col: Style last column differently
        header_background: Header row background color (hex, e.g., "#0066CC")
        header_text_color: Header row text color (hex, e.g., "#FFFFFF")
        cell_borders: Show cell borders (default True)
        border_color: Border color (default "#000000")

    Example:
        >>> # Simple table with default styling
        >>> add_table(pres, 4, data=[
        ...     ["Region", "Q3", "Q4"],
        ...     ["North", 100, 120],
        ...     ["South", 80, 95],
        ... ])

        >>> # Custom header styling
        >>> add_table(pres, 5, data=[...],
        ...     header_background="#003366",
        ...     header_text_color="#FFFFFF",
        ...     banded_rows=True
        ... )

        >>> # Table without borders
        >>> add_table(pres, 6, data=[...], cell_borders=False)
    """
    slide = presentation.get_slide(slide_number)

    # Build styled table through core
    slide.add_table(
        data,
        left=left,
        top=top,
        width=width,
        height=height,
        placeholder=placeholder,
        header_row=header_row,
        banded_rows=banded_rows,
        first_col=first_col,
        last_col=last_col,
        header_background=header_background,
        header_text_color=header_text_color,
        cell_borders=cell_borders,
        border_color=border_color,
    )


def update_table_cell(
    presentation: Presentation,
    slide_number: int,
    table_index: int,
    row: int,
    col: int,
    value: Any,
) -> None:
    """Update a single cell in a table.

    Args:
        presentation: The presentation to modify
        slide_number: The slide number (1-indexed)
        table_index: Index of the table on the slide (0-indexed)
        row: Row index (0-indexed)
        col: Column index (0-indexed)
        value: New cell value

    Example:
        >>> update_table_cell(pres, 4, table_index=0, row=1, col=2, value=125)
    """
    slide = presentation.get_slide(slide_number)

    # Find tables on the slide
    tables = [s for s in slide.shapes if isinstance(s, Table)]

    if table_index >= len(tables):
        raise ValueError(
            f"Table index {table_index} out of range. "
            f"Slide has {len(tables)} table(s)."
        )

    table = tables[table_index]

    if row >= len(table.rows):
        raise ValueError(
            f"Row {row} out of range. Table has {len(table.rows)} rows."
        )

    if col >= len(table.rows[row]):
        raise ValueError(
            f"Column {col} out of range. Table has {len(table.rows[row])} columns."
        )

    # Update the cell
    table.rows[row][col].text = str(value) if value is not None else ""

    # Save changes
    from ..oxml.slide import update_slide_in_package

    update_slide_in_package(
        presentation._package,
        slide_number,
        slide._part,
    )


def style_table_cell(
    presentation: Presentation,
    slide_number: int,
    table_index: int,
    row: int,
    col: int,
    *,
    background: str | None = None,
    bold: bool | None = None,
    font_size: int | None = None,
    color: str | None = None,
    border_top: bool | str | None = None,
    border_bottom: bool | str | None = None,
    border_left: bool | str | None = None,
    border_right: bool | str | None = None,
    vertical_align: str | None = None,
) -> None:
    """Style a specific cell in a table.

    Args:
        presentation: The presentation to modify
        slide_number: The slide number (1-indexed)
        table_index: Index of the table on the slide (0-indexed)
        row: Row index (0-indexed)
        col: Column index (0-indexed)
        background: Background color as hex (e.g., "#FFFF00")
        bold: Make text bold
        font_size: Font size in points
        color: Text color as hex (e.g., "#FF0000")
        border_top: Top border (True for default, hex color, or False for none)
        border_bottom: Bottom border
        border_left: Left border
        border_right: Right border
        vertical_align: Vertical alignment ("top", "center", "bottom")

    Example:
        >>> # Highlight a cell yellow with bold text
        >>> style_table_cell(pres, 4, 0, row=1, col=2,
        ...     background="#FFFF00", bold=True)

        >>> # Remove borders from a cell
        >>> style_table_cell(pres, 4, 0, row=0, col=0,
        ...     border_top=False, border_bottom=False)
    """
    slide = presentation.get_slide(slide_number)
    tables = [s for s in slide.shapes if isinstance(s, Table)]

    if table_index >= len(tables):
        raise ValueError(
            f"Table index {table_index} out of range. "
            f"Slide has {len(tables)} table(s)."
        )

    table = tables[table_index]

    if row >= len(table.rows):
        raise ValueError(
            f"Row {row} out of range. Table has {len(table.rows)} rows."
        )

    if col >= len(table.rows[row]):
        raise ValueError(
            f"Column {col} out of range. Table has {len(table.rows[row])} columns."
        )

    cell = table.rows[row][col]

    # Initialize style if needed
    if cell.style is None:
        cell.style = CellStyle()

    # Apply text formatting
    if bold is not None:
        cell.bold = bold
    if font_size is not None:
        cell.font_size = font_size * 100  # Convert to centipoints
    if color is not None:
        cell.color = color.lstrip("#")

    # Apply background
    if background is not None:
        cell.style.background_color = background.lstrip("#")

    # Apply vertical alignment
    if vertical_align is not None:
        align_map = {"top": "t", "center": "ctr", "bottom": "b"}
        cell.style.vertical_align = align_map.get(vertical_align.lower(), "ctr")

    # Helper to create border style
    def make_border(value):
        if value is False:
            return BorderStyle(style="none")
        elif value is True:
            return BorderStyle()  # Default black border
        elif isinstance(value, str):
            return BorderStyle(color=value.lstrip("#"))
        return None

    # Apply borders
    if border_top is not None:
        cell.style.border_top = make_border(border_top)
    if border_bottom is not None:
        cell.style.border_bottom = make_border(border_bottom)
    if border_left is not None:
        cell.style.border_left = make_border(border_left)
    if border_right is not None:
        cell.style.border_right = make_border(border_right)

    # Save changes
    from ..oxml.slide import update_slide_in_package

    update_slide_in_package(
        presentation._package,
        slide_number,
        slide._part,
    )


def merge_table_cells(
    presentation: Presentation,
    slide_number: int,
    table_index: int,
    start_row: int,
    start_col: int,
    end_row: int,
    end_col: int,
) -> None:
    """Merge a range of cells in a table.

    The merged cell will contain the text from the top-left cell.
    Other cells in the range will be cleared.

    Args:
        presentation: The presentation to modify
        slide_number: The slide number (1-indexed)
        table_index: Index of the table on the slide (0-indexed)
        start_row: Starting row (0-indexed, inclusive)
        start_col: Starting column (0-indexed, inclusive)
        end_row: Ending row (0-indexed, inclusive)
        end_col: Ending column (0-indexed, inclusive)

    Example:
        >>> # Merge cells (0,0) through (0,2) - merge first 3 columns of header
        >>> merge_table_cells(pres, 4, 0,
        ...     start_row=0, start_col=0, end_row=0, end_col=2)

        >>> # Merge a 2x2 block
        >>> merge_table_cells(pres, 4, 0,
        ...     start_row=1, start_col=0, end_row=2, end_col=1)
    """
    slide = presentation.get_slide(slide_number)
    tables = [s for s in slide.shapes if isinstance(s, Table)]

    if table_index >= len(tables):
        raise ValueError(
            f"Table index {table_index} out of range. "
            f"Slide has {len(tables)} table(s)."
        )

    table = tables[table_index]

    # Validate range
    if start_row < 0 or end_row >= len(table.rows):
        raise ValueError(f"Row range ({start_row}-{end_row}) out of bounds")
    if start_col < 0 or end_col >= len(table.rows[0]):
        raise ValueError(f"Column range ({start_col}-{end_col}) out of bounds")
    if start_row > end_row or start_col > end_col:
        raise ValueError("Start must be less than or equal to end")

    row_span = end_row - start_row + 1
    col_span = end_col - start_col + 1

    # Set the origin cell
    origin_cell = table.rows[start_row][start_col]
    origin_cell.row_span = row_span
    origin_cell.col_span = col_span
    origin_cell.is_merge_origin = True

    # Mark other cells as merged (they need to exist but will be empty)
    for r in range(start_row, end_row + 1):
        for c in range(start_col, end_col + 1):
            if r == start_row and c == start_col:
                continue  # Skip origin
            cell = table.rows[r][c]
            cell.text = ""
            cell.is_merge_origin = False
            # Horizontal merge marker
            if r == start_row:
                cell.col_span = 0
            # Vertical merge marker
            if c == start_col:
                cell.row_span = 0

    # Save changes
    from ..oxml.slide import update_slide_in_package

    update_slide_in_package(
        presentation._package,
        slide_number,
        slide._part,
    )


def add_image(
    presentation: Presentation,
    slide_number: int,
    image_path: str | Path,
    *,
    placeholder: str | None = None,
    left: str | int | None = None,
    top: str | int | None = None,
    width: str | int | None = None,
    height: str | int | None = None,
    rotation: int = 0,
    crop: dict | None = None,
    shadow: bool = False,
    reflection: bool = False,
    brightness: int = 0,
    contrast: int = 0,
) -> None:
    """Add an image to a slide.

    Args:
        presentation: The presentation to modify
        slide_number: The slide number (1-indexed)
        image_path: Path to the image file (PNG, JPEG, GIF, etc.)
        placeholder: Placeholder to fill (e.g., "content", "picture").
                    If provided, image will use placeholder's position.
        left: Left position (e.g., "1in", "2.5cm") - ignored if placeholder set
        top: Top position - ignored if placeholder set
        width: Image width - ignored if placeholder set
        height: Image height - ignored if placeholder set
        rotation: Rotation angle in degrees (0-360)
        crop: Crop the image as dict with "left", "top", "right", "bottom"
              percentages (0-100). E.g., {"left": 10, "right": 10} crops
              10% from each side.
        shadow: Add drop shadow effect
        reflection: Add reflection effect
        brightness: Brightness adjustment (-100 to 100)
        contrast: Contrast adjustment (-100 to 100)

    Example:
        >>> # Simple image
        >>> add_image(pres, 3, "chart.png", placeholder="content")

        >>> # Image with positioning
        >>> add_image(pres, 3, "logo.png",
        ...     left="7in", top="0.5in", width="2in")

        >>> # Image with effects
        >>> add_image(pres, 3, "photo.jpg",
        ...     left="1in", top="1in", width="4in", height="3in",
        ...     shadow=True, brightness=10)

        >>> # Cropped image
        >>> add_image(pres, 3, "wide_photo.jpg",
        ...     left="1in", top="1in", width="4in", height="3in",
        ...     crop={"left": 20, "right": 20})  # Crop 20% from each side
    """
    image_path = Path(image_path)

    if not image_path.exists():
        raise FileNotFoundError(f"Image file not found: {image_path}")

    # Read image data
    with open(image_path, "rb") as f:
        image_data = f.read()

    # Determine content type
    ext = image_path.suffix.lower()
    content_types = {
        ".png": CONTENT_TYPE.PNG,
        ".jpg": CONTENT_TYPE.JPEG,
        ".jpeg": CONTENT_TYPE.JPEG,
        ".gif": CONTENT_TYPE.GIF,
        ".bmp": CONTENT_TYPE.BMP,
        ".tiff": CONTENT_TYPE.TIFF,
        ".tif": CONTENT_TYPE.TIFF,
        ".emf": CONTENT_TYPE.EMF,
        ".wmf": CONTENT_TYPE.WMF,
    }

    content_type = content_types.get(ext)
    if content_type is None:
        raise ValueError(f"Unsupported image format: {ext}")

    # Add image to package
    pkg = presentation._package

    # Find next available image number
    existing_images = [
        name for name, _ in pkg.iter_parts()
        if name.startswith("ppt/media/image")
    ]
    image_num = len(existing_images) + 1
    image_part_name = f"ppt/media/image{image_num}{ext}"

    pkg.set_part(image_part_name, image_data, content_type)

    # Get slide and determine position
    slide = presentation.get_slide(slide_number)

    if placeholder:
        ph_shape = slide._find_placeholder(placeholder)
        position = ph_shape.position
    elif left is not None and top is not None:
        position = Position(
            x=int(parse_length(left)),
            y=int(parse_length(top)),
            cx=int(parse_length(width or "4in")),
            cy=int(parse_length(height or "3in")),
        )
    else:
        # Default position (centered)
        position = Position(
            x=presentation.slide_width // 4,
            y=presentation.slide_height // 4,
            cx=presentation.slide_width // 2,
            cy=presentation.slide_height // 2,
        )

    # Add relationship to slide
    slide_refs = presentation._presentation.get_slide_refs()
    slide_ref = slide_refs[slide_number - 1]
    pres_rels = pkg.get_part_rels("ppt/presentation.xml")
    rel = pres_rels.get(slide_ref.r_id)

    if rel.target.startswith("/"):
        slide_path = rel.target.lstrip("/")
    else:
        slide_path = f"ppt/{rel.target}"

    slide_rels = pkg.get_part_rels(slide_path)
    r_id = slide_rels.add(
        rel_type=REL_TYPE.IMAGE,
        target=f"../media/image{image_num}{ext}",
    )
    pkg.set_part_rels(slide_path, slide_rels)

    # Build crop rect if provided
    crop_rect = None
    if crop:
        crop_rect = CropRect(
            left=int(crop.get("left", 0) * 1000),
            top=int(crop.get("top", 0) * 1000),
            right=int(crop.get("right", 0) * 1000),
            bottom=int(crop.get("bottom", 0) * 1000),
        )

    # Build effects if any specified
    effects = None
    if shadow or reflection or brightness != 0 or contrast != 0:
        effects = PictureEffects(
            shadow=shadow,
            reflection=reflection,
            brightness=brightness,
            contrast=contrast,
        )

    # Create picture
    pic = Picture(
        id=slide._part.shape_tree._next_id,
        name=f"Picture {image_num}",
        position=position,
        r_embed=r_id,
        rotation=rotation,
        crop=crop_rect,
        effects=effects,
    )

    slide._part.add_picture(pic)
    slide._save()


def crop_image(
    presentation: Presentation,
    slide_number: int,
    image_index: int,
    *,
    left: int = 0,
    top: int = 0,
    right: int = 0,
    bottom: int = 0,
) -> None:
    """Crop an existing image on a slide.

    Args:
        presentation: The presentation to modify
        slide_number: The slide number (1-indexed)
        image_index: Index of the image on the slide (0-indexed)
        left: Percentage to crop from left (0-100)
        top: Percentage to crop from top (0-100)
        right: Percentage to crop from right (0-100)
        bottom: Percentage to crop from bottom (0-100)

    Example:
        >>> # Crop 10% from each side
        >>> crop_image(pres, 1, 0, left=10, top=10, right=10, bottom=10)
    """
    slide = presentation.get_slide(slide_number)

    # Find images on the slide
    images = [s for s in slide.shapes if isinstance(s, Picture)]

    if image_index >= len(images):
        raise ValueError(
            f"Image index {image_index} out of range. "
            f"Slide has {len(images)} image(s)."
        )

    image = images[image_index]

    # Apply crop
    image.crop = CropRect(
        left=int(left * 1000),
        top=int(top * 1000),
        right=int(right * 1000),
        bottom=int(bottom * 1000),
    )

    slide._save()


def rotate_image(
    presentation: Presentation,
    slide_number: int,
    image_index: int,
    angle: int,
) -> None:
    """Rotate an existing image on a slide.

    Args:
        presentation: The presentation to modify
        slide_number: The slide number (1-indexed)
        image_index: Index of the image on the slide (0-indexed)
        angle: Rotation angle in degrees (positive = clockwise)

    Example:
        >>> # Rotate image 45 degrees
        >>> rotate_image(pres, 1, 0, 45)
    """
    slide = presentation.get_slide(slide_number)

    # Find images on the slide
    images = [s for s in slide.shapes if isinstance(s, Picture)]

    if image_index >= len(images):
        raise ValueError(
            f"Image index {image_index} out of range. "
            f"Slide has {len(images)} image(s)."
        )

    image = images[image_index]
    image.rotation = angle

    slide._save()


def flip_image(
    presentation: Presentation,
    slide_number: int,
    image_index: int,
    *,
    horizontal: bool = False,
    vertical: bool = False,
) -> None:
    """Flip an existing image on a slide.

    Args:
        presentation: The presentation to modify
        slide_number: The slide number (1-indexed)
        image_index: Index of the image on the slide (0-indexed)
        horizontal: Flip horizontally
        vertical: Flip vertically

    Example:
        >>> # Mirror image horizontally
        >>> flip_image(pres, 1, 0, horizontal=True)
    """
    slide = presentation.get_slide(slide_number)

    # Find images on the slide
    images = [s for s in slide.shapes if isinstance(s, Picture)]

    if image_index >= len(images):
        raise ValueError(
            f"Image index {image_index} out of range. "
            f"Slide has {len(images)} image(s)."
        )

    image = images[image_index]
    if horizontal:
        image.flip_h = not image.flip_h
    if vertical:
        image.flip_v = not image.flip_v

    slide._save()


def add_chart(
    presentation: Presentation,
    slide_number: int,
    chart_type: str,
    categories: list[str],
    series: list[dict],
    *,
    title: str | None = None,
    placeholder: str | None = None,
    left: str | int | None = None,
    top: str | int | None = None,
    width: str | int | None = None,
    height: str | int | None = None,
    stacked: bool = False,
    percent_stacked: bool = False,
    legend: Literal["right", "left", "top", "bottom", "none"] = "right",
    data_labels: bool = False,
    markers: bool = True,
    smooth: bool = False,
) -> None:
    """Add a chart to a slide.

    Charts use theme colors by default for series, ensuring they match
    the template's branding. You can override with explicit colors.

    Args:
        presentation: The presentation to modify
        slide_number: The slide number (1-indexed)
        chart_type: Type of chart. Supported types:
            - "column" / "bar" (vertical/horizontal bar charts)
            - "stacked_column" / "stacked_bar"
            - "line" / "line_markers"
            - "pie" / "doughnut"
            - "area" / "stacked_area"
            - "scatter" / "scatter_lines" / "scatter_smooth"
        categories: Category labels (x-axis labels for most charts)
        series: List of data series, each a dict with:
            - "name": Series name (shown in legend)
            - "values": List of numeric values
            - "color": Optional hex color ("#FF0000") or theme color ("accent1")
        title: Optional chart title
        placeholder: Placeholder to fill (e.g., "content", "chart").
                    If provided, chart will use placeholder's position.
        left: Left position (e.g., "1in", "2.5cm") - ignored if placeholder set
        top: Top position - ignored if placeholder set
        width: Chart width - ignored if placeholder set
        height: Chart height - ignored if placeholder set
        stacked: Stack series (for bar/column/area charts)
        percent_stacked: Stack to 100% (overrides stacked)
        legend: Legend position ("right", "left", "top", "bottom", "none")
        data_labels: Show data values on chart
        markers: Show markers on line/scatter charts
        smooth: Use smooth lines on line/scatter charts

    Example:
        >>> # Simple column chart using theme colors
        >>> add_chart(pres, 2, "column",
        ...     categories=["Q1", "Q2", "Q3", "Q4"],
        ...     series=[
        ...         {"name": "2023", "values": [100, 120, 140, 160]},
        ...         {"name": "2024", "values": [110, 135, 155, 180]},
        ...     ],
        ...     title="Quarterly Revenue"
        ... )

        >>> # Pie chart in content placeholder
        >>> add_chart(pres, 3, "pie",
        ...     categories=["North", "South", "East", "West"],
        ...     series=[{"name": "Sales", "values": [30, 25, 20, 25]}],
        ...     placeholder="content",
        ...     data_labels=True
        ... )

        >>> # Scatter chart with custom colors
        >>> add_chart(pres, 4, "scatter",
        ...     categories=[1, 2, 3, 4, 5],  # x values
        ...     series=[
        ...         {"name": "Actual", "values": [2.1, 4.2, 5.8, 8.1, 9.9],
        ...          "color": "accent1"},
        ...         {"name": "Target", "values": [2, 4, 6, 8, 10],
        ...          "color": "#666666"},
        ...     ],
        ...     markers=True, smooth=True
        ... )
    """
    slide = presentation.get_slide(slide_number)
    pkg = presentation._package

    # Build chart data
    chart_series = []
    for s in series:
        chart_series.append(
            ChartSeries(
                name=s.get("name", ""),
                values=s.get("values", []),
                color=s.get("color"),
                categories=s.get("categories"),  # For scatter with per-series x
            )
        )

    chart_data = ChartData(
        categories=[str(c) for c in categories],
        series=chart_series,
    )

    # Build chart properties
    chart_props = ChartProperties(
        title=title,
        legend=legend,
        data_labels=data_labels,
        stacked=stacked,
        percent_stacked=percent_stacked,
        smooth=smooth,
        markers=markers,
        bar_direction="bar" if "bar" in chart_type.lower() else "col",
    )

    # Create chart part
    chart_part = ChartPart.new(chart_type, chart_data, chart_props)
    chart_xml = chart_part.to_xml()

    # Find next available chart number
    existing_charts = [
        name for name, _ in pkg.iter_parts()
        if name.startswith("ppt/charts/chart")
    ]
    chart_num = len(existing_charts) + 1
    chart_part_name = f"ppt/charts/chart{chart_num}.xml"

    # Add chart part to package
    pkg.set_part(chart_part_name, chart_xml, CONTENT_TYPE.CHART)

    # Determine position
    if placeholder:
        ph_shape = slide._find_placeholder(placeholder)
        position = ph_shape.position
    elif left is not None and top is not None:
        position = Position(
            x=int(parse_length(left)),
            y=int(parse_length(top)),
            cx=int(parse_length(width or "6in")),
            cy=int(parse_length(height or "4in")),
        )
    else:
        # Default position (centered, reasonable size)
        position = Position(
            x=presentation.slide_width // 6,
            y=presentation.slide_height // 4,
            cx=int(presentation.slide_width * 0.67),
            cy=int(presentation.slide_height * 0.5),
        )

    # Get slide path and add relationship
    slide_refs = presentation._presentation.get_slide_refs()
    slide_ref = slide_refs[slide_number - 1]
    pres_rels = pkg.get_part_rels("ppt/presentation.xml")
    rel = pres_rels.get(slide_ref.r_id)

    if rel.target.startswith("/"):
        slide_path = rel.target.lstrip("/")
    else:
        slide_path = f"ppt/{rel.target}"

    slide_rels = pkg.get_part_rels(slide_path)
    r_id = slide_rels.add(
        rel_type=REL_TYPE.CHART,
        target=f"../charts/chart{chart_num}.xml",
    )
    pkg.set_part_rels(slide_path, slide_rels)

    # Create chart shape
    chart_shape = Chart(
        id=slide._part.shape_tree._next_id,
        name=f"Chart {chart_num}",
        position=position,
        r_embed=r_id,
    )

    slide._part.add_chart(chart_shape)
    slide._save()


def update_chart_data(
    presentation: Presentation,
    slide_number: int,
    chart_index: int,
    categories: list[str],
    series: list[dict],
) -> None:
    """Update the data in an existing chart.

    Args:
        presentation: The presentation to modify
        slide_number: The slide number (1-indexed)
        chart_index: Index of the chart on the slide (0-indexed)
        categories: New category labels
        series: New data series (same format as add_chart)

    Example:
        >>> # Update first chart on slide 2
        >>> update_chart_data(pres, 2, 0,
        ...     categories=["Q1", "Q2", "Q3", "Q4"],
        ...     series=[
        ...         {"name": "Updated", "values": [150, 180, 200, 220]},
        ...     ]
        ... )
    """
    slide = presentation.get_slide(slide_number)
    pkg = presentation._package

    # Find charts on the slide
    charts = [s for s in slide.shapes if isinstance(s, Chart)]

    if chart_index >= len(charts):
        raise ValueError(
            f"Chart index {chart_index} out of range. "
            f"Slide has {len(charts)} chart(s)."
        )

    chart_shape = charts[chart_index]

    # Get slide path and chart relationship
    slide_refs = presentation._presentation.get_slide_refs()
    slide_ref = slide_refs[slide_number - 1]
    pres_rels = pkg.get_part_rels("ppt/presentation.xml")
    rel = pres_rels.get(slide_ref.r_id)

    if rel.target.startswith("/"):
        slide_path = rel.target.lstrip("/")
    else:
        slide_path = f"ppt/{rel.target}"

    slide_rels = pkg.get_part_rels(slide_path)
    chart_rel = slide_rels.get(chart_shape.r_embed)

    if chart_rel is None:
        raise ValueError("Could not find chart relationship")

    # Resolve chart path
    if chart_rel.target.startswith(".."):
        # Relative path from slides/
        chart_path = f"ppt/charts/{chart_rel.target.split('/')[-1]}"
    else:
        chart_path = chart_rel.target.lstrip("/")

    # Get existing chart XML
    chart_xml = pkg.get_part(chart_path)
    if chart_xml is None:
        raise ValueError("Could not find chart part")

    # Parse existing chart to get its type and properties
    existing_chart = ChartPart.from_xml(chart_xml)

    # Build new chart data
    chart_series = []
    for s in series:
        chart_series.append(
            ChartSeries(
                name=s.get("name", ""),
                values=s.get("values", []),
                color=s.get("color"),
            )
        )

    # Create new chart with same type but new data
    new_chart_data = ChartData(
        categories=[str(c) for c in categories],
        series=chart_series,
    )

    # Preserve chart type
    new_chart = ChartPart.new(existing_chart.chart_type, new_chart_data, existing_chart.props)

    # Update the chart part
    pkg.set_part(chart_path, new_chart.to_xml(), CONTENT_TYPE.CHART)
