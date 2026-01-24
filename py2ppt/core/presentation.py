"""High-level Presentation class."""

from __future__ import annotations

from pathlib import Path
from typing import BinaryIO

from ..oxml.layout import (
    LayoutInfo,
    SlideLayoutPart,
    get_layout_by_index,
    get_layout_by_name,
    get_layout_info_list,
)
from ..oxml.ns import CONTENT_TYPE, REL_TYPE
from ..oxml.package import Package, create_blank_package
from ..oxml.presentation import (
    PresentationPart,
    get_presentation_part,
    setup_presentation_part,
)
from ..oxml.slide import (
    SlidePart,
    add_slide_to_package,
    get_slide_part,
    remove_slide_from_package,
)
from ..oxml.theme import ThemePart, get_theme_part
from ..utils.errors import (
    InvalidTemplateError,
    LayoutNotFoundError,
    SlideNotFoundError,
    find_similar,
)
from .slide import Slide


class Presentation:
    """High-level presentation object.

    This class wraps the low-level package and provides a clean API
    for working with presentations.
    """

    def __init__(self, package: Package) -> None:
        """Initialize from a package.

        Use the class methods `new()`, `open()`, or `from_template()`
        to create instances.
        """
        self._package = package
        self._pres_part: PresentationPart | None = None
        self._theme: ThemePart | None = None
        self._layouts: list[LayoutInfo] | None = None
        self._dirty = False

    @property
    def package(self) -> Package:
        """Get the underlying package."""
        return self._package

    @property
    def _presentation(self) -> PresentationPart:
        """Get the presentation part (lazy loaded)."""
        if self._pres_part is None:
            self._pres_part = get_presentation_part(self._package)
            if self._pres_part is None:
                raise InvalidTemplateError("No presentation part found in package")
        return self._pres_part

    @property
    def slide_count(self) -> int:
        """Get the number of slides."""
        return len(self._presentation.get_slide_refs())

    @property
    def slide_width(self) -> int:
        """Get slide width in EMUs."""
        return self._presentation.get_slide_size()[0]

    @property
    def slide_height(self) -> int:
        """Get slide height in EMUs."""
        return self._presentation.get_slide_size()[1]

    def get_slide(self, slide_number: int) -> Slide:
        """Get a slide by number (1-indexed).

        Args:
            slide_number: Slide number (1-indexed)

        Returns:
            Slide object

        Raises:
            SlideNotFoundError: If slide number is out of range
        """
        if slide_number < 1 or slide_number > self.slide_count:
            raise SlideNotFoundError(slide_number, self.slide_count)

        slide_part = get_slide_part(self._package, slide_number)
        if slide_part is None:
            raise SlideNotFoundError(slide_number, self.slide_count)

        return Slide(slide_part, slide_number, self)

    def get_layouts(self) -> list[LayoutInfo]:
        """Get information about all available layouts."""
        if self._layouts is None:
            self._layouts = get_layout_info_list(self._package)
        return self._layouts

    def get_layout_names(self) -> list[str]:
        """Get list of layout names."""
        return [layout.name for layout in self.get_layouts()]

    def get_theme_colors(self) -> dict[str, str]:
        """Get theme colors as name -> hex color dict."""
        if self._theme is None:
            self._theme = get_theme_part(self._package)
        if self._theme:
            colors = self._theme.get_colors()
            return {name: f"#{rgb}" for name, rgb in colors.items()}
        return {}

    def get_theme_fonts(self) -> dict[str, str]:
        """Get theme fonts as role -> font name dict."""
        if self._theme is None:
            self._theme = get_theme_part(self._package)
        if self._theme:
            fonts = self._theme.get_fonts()
            return {
                "heading": fonts.major_font.typeface,
                "body": fonts.minor_font.typeface,
            }
        return {"heading": "Calibri Light", "body": "Calibri"}

    def add_slide(
        self,
        layout: str | int = 0,
        position: int | None = None,
    ) -> Slide:
        """Add a new slide.

        Args:
            layout: Layout name (fuzzy matched) or index (0-indexed)
            position: Insert position (1-indexed). None = append at end.

        Returns:
            The new Slide object

        Raises:
            LayoutNotFoundError: If layout name not found
        """
        # Find layout
        if isinstance(layout, str):
            result = get_layout_by_name(self._package, layout, fuzzy=True)
            if result is None:
                layout_names = self.get_layout_names()
                suggestion = find_similar(layout, layout_names)
                raise LayoutNotFoundError(layout, layout_names, suggestion)
            layout_part, layout_idx = result
        else:
            layout_idx = layout
            layout_part = get_layout_by_index(self._package, layout_idx)
            if layout_part is None:
                raise LayoutNotFoundError(
                    str(layout),
                    available=[f"{i}: {name}" for i, name in enumerate(self.get_layout_names())],
                )

        # Create new slide from layout
        slide_part = SlidePart.new()

        # Copy placeholders from layout to slide
        for ph in layout_part.get_placeholders():
            from ..oxml.shapes import PlaceholderInfo, Position, Shape, TextFrame

            shape = Shape(
                id=slide_part.shape_tree._next_id,
                name=ph.name,
                position=Position(
                    x=ph.position.x,
                    y=ph.position.y,
                    cx=ph.position.cx,
                    cy=ph.position.cy,
                ),
                placeholder=PlaceholderInfo(
                    type=ph.type,
                    idx=ph.idx,
                ),
                text_frame=TextFrame(),
                preset_geometry="rect",
            )
            slide_part.add_shape(shape)

        # Add to package
        # Position is 0-indexed for internal use
        insert_pos = None if position is None else position - 1
        slide_num = add_slide_to_package(
            self._package,
            slide_part,
            f"rId{layout_idx + 1}",
            insert_pos,
        )

        # Invalidate cache - the package presentation.xml has been updated
        self._pres_part = None
        self._dirty = True

        return self.get_slide(slide_num)

    def delete_slide(self, slide_number: int) -> bool:
        """Delete a slide.

        Args:
            slide_number: Slide number to delete (1-indexed)

        Returns:
            True if deleted, False if slide not found
        """
        if slide_number < 1 or slide_number > self.slide_count:
            return False

        result = remove_slide_from_package(self._package, slide_number)
        if result:
            # Invalidate cache
            self._pres_part = None
            self._dirty = True
        return result

    def duplicate_slide(self, slide_number: int) -> Slide:
        """Duplicate a slide.

        Args:
            slide_number: Slide number to duplicate (1-indexed)

        Returns:
            The new (duplicated) Slide object

        Raises:
            SlideNotFoundError: If slide number is out of range
        """
        if slide_number < 1 or slide_number > self.slide_count:
            raise SlideNotFoundError(slide_number, self.slide_count)

        # Get the slide to duplicate (kept for future content copying)
        _original = self.get_slide(slide_number)

        # Add new slide with same layout
        # For now, use first layout - proper implementation would copy layout info
        new_slide = self.add_slide(layout=0, position=slide_number + 1)

        # Copy content
        # TODO: Implement full content copying
        self._dirty = True

        return new_slide

    def reorder_slides(self, new_order: list[int]) -> None:
        """Reorder slides.

        Args:
            new_order: New order as list of slide numbers (1-indexed)
                       e.g., [2, 1, 3] moves slide 2 to first position

        Raises:
            ValueError: If new_order is invalid
        """
        if sorted(new_order) != list(range(1, self.slide_count + 1)):
            raise ValueError(
                f"new_order must contain all slide numbers 1-{self.slide_count}"
            )

        # Get current slide refs
        slide_refs = self._presentation.get_slide_refs()

        # Build new order of rIds
        new_r_ids = [slide_refs[num - 1].r_id for num in new_order]

        # Reorder
        self._presentation.reorder_slides(new_r_ids)

        # Update package
        self._package.set_part(
            "ppt/presentation.xml",
            self._presentation.to_xml(),
            CONTENT_TYPE.PRESENTATION,
        )

        self._dirty = True

    def save(self, path: str | Path | BinaryIO) -> None:
        """Save the presentation.

        Args:
            path: File path or file-like object
        """
        self._package.save(path)
        self._dirty = False

    def to_bytes(self) -> bytes:
        """Get presentation as bytes."""
        return self._package.to_bytes()

    # === Class Methods ===

    @classmethod
    def new(cls) -> Presentation:
        """Create a new blank presentation."""
        pkg = create_blank_package()

        # Create presentation part
        pres_part = PresentationPart.new()

        # Create theme
        theme = ThemePart.new()
        pkg.set_part("ppt/theme/theme1.xml", theme.to_xml(), CONTENT_TYPE.THEME)

        # Create minimal master
        from ..oxml.master import create_minimal_master

        master = create_minimal_master()
        pkg.set_part(
            "ppt/slideMasters/slideMaster1.xml",
            master.to_xml(),
            CONTENT_TYPE.SLIDE_MASTER,
        )

        # Create basic layouts
        layouts = _create_default_layouts()
        for i, layout in enumerate(layouts, 1):
            pkg.set_part(
                f"ppt/slideLayouts/slideLayout{i}.xml",
                layout.to_xml(),
                CONTENT_TYPE.SLIDE_LAYOUT,
            )

            # Add layout relationship to master
            master_rels = pkg.get_part_rels("ppt/slideMasters/slideMaster1.xml")
            master_rels.add(
                rel_type=REL_TYPE.SLIDE_LAYOUT,
                target=f"../slideLayouts/slideLayout{i}.xml",
            )
            pkg.set_part_rels("ppt/slideMasters/slideMaster1.xml", master_rels)

        # Add master relationship to presentation
        pres_rels = pkg.get_part_rels("ppt/presentation.xml")
        pres_rels.add(
            rel_type=REL_TYPE.SLIDE_MASTER,
            target="slideMasters/slideMaster1.xml",
        )
        pres_rels.add(
            rel_type=REL_TYPE.THEME,
            target="theme/theme1.xml",
        )
        pkg.set_part_rels("ppt/presentation.xml", pres_rels)

        # Add master reference to presentation part
        pres_part._element.find(".//p:sldMasterIdLst", namespaces={"p": "http://schemas.openxmlformats.org/presentationml/2006/main"})
        from lxml import etree

        from ..oxml.ns import qn

        master_lst = pres_part._element.find(qn("p:sldMasterIdLst"))
        if master_lst is not None:
            master_id = etree.SubElement(master_lst, qn("p:sldMasterId"))
            master_id.set("id", "2147483648")
            master_id.set(qn("r:id"), "rId1")

        setup_presentation_part(pkg, pres_part)

        return cls(pkg)

    @classmethod
    def open(cls, path: str | Path | BinaryIO) -> Presentation:
        """Open an existing presentation.

        Args:
            path: File path or file-like object

        Returns:
            Presentation object
        """
        pkg = Package.open(path)
        return cls(pkg)

    @classmethod
    def from_template(cls, template_path: str | Path) -> Presentation:
        """Create a new presentation from a template.

        The template is opened and all slides are removed, leaving
        only the layouts, masters, and theme.

        Args:
            template_path: Path to template file

        Returns:
            New Presentation object
        """
        pres = cls.open(template_path)

        # Remove all existing slides
        while pres.slide_count > 0:
            pres.delete_slide(1)

        return pres


def _create_default_layouts() -> list[SlideLayoutPart]:
    """Create default slide layouts."""
    from ..oxml.shapes import (
        PlaceholderInfo,
        Position,
        Shape,
        TextFrame,
    )

    layouts = []

    # Layout 1: Title Slide
    layout1 = SlideLayoutPart.new("Title Slide")
    # Title placeholder (centered)
    title_shape = Shape(
        id=2,
        name="Title",
        position=Position(x=685800, y=2130425, cx=7772400, cy=1470025),
        placeholder=PlaceholderInfo(type="ctrTitle"),
        text_frame=TextFrame(),
        preset_geometry="rect",
    )
    layout1.shape_tree.add_shape(title_shape)
    # Subtitle placeholder
    subtitle_shape = Shape(
        id=3,
        name="Subtitle",
        position=Position(x=1371600, y=3886200, cx=6400800, cy=1752600),
        placeholder=PlaceholderInfo(type="subTitle", idx=1),
        text_frame=TextFrame(),
        preset_geometry="rect",
    )
    layout1.shape_tree.add_shape(subtitle_shape)
    layouts.append(layout1)

    # Layout 2: Title and Content
    layout2 = SlideLayoutPart.new("Title and Content")
    title_shape = Shape(
        id=2,
        name="Title",
        position=Position(x=457200, y=274638, cx=8229600, cy=1143000),
        placeholder=PlaceholderInfo(type="title"),
        text_frame=TextFrame(),
        preset_geometry="rect",
    )
    layout2.shape_tree.add_shape(title_shape)
    body_shape = Shape(
        id=3,
        name="Content Placeholder",
        position=Position(x=457200, y=1600200, cx=8229600, cy=4525963),
        placeholder=PlaceholderInfo(type="body", idx=1),
        text_frame=TextFrame(),
        preset_geometry="rect",
    )
    layout2.shape_tree.add_shape(body_shape)
    layouts.append(layout2)

    # Layout 3: Section Header
    layout3 = SlideLayoutPart.new("Section Header")
    title_shape = Shape(
        id=2,
        name="Title",
        position=Position(x=722313, y=4406900, cx=7772400, cy=1362075),
        placeholder=PlaceholderInfo(type="title"),
        text_frame=TextFrame(),
        preset_geometry="rect",
    )
    layout3.shape_tree.add_shape(title_shape)
    layouts.append(layout3)

    # Layout 4: Two Content
    layout4 = SlideLayoutPart.new("Two Content")
    title_shape = Shape(
        id=2,
        name="Title",
        position=Position(x=457200, y=274638, cx=8229600, cy=1143000),
        placeholder=PlaceholderInfo(type="title"),
        text_frame=TextFrame(),
        preset_geometry="rect",
    )
    layout4.shape_tree.add_shape(title_shape)
    left_shape = Shape(
        id=3,
        name="Content Placeholder",
        position=Position(x=457200, y=1600200, cx=4038600, cy=4525963),
        placeholder=PlaceholderInfo(type="body", idx=1),
        text_frame=TextFrame(),
        preset_geometry="rect",
    )
    layout4.shape_tree.add_shape(left_shape)
    right_shape = Shape(
        id=4,
        name="Content Placeholder 2",
        position=Position(x=4648200, y=1600200, cx=4038600, cy=4525963),
        placeholder=PlaceholderInfo(type="body", idx=2),
        text_frame=TextFrame(),
        preset_geometry="rect",
    )
    layout4.shape_tree.add_shape(right_shape)
    layouts.append(layout4)

    # Layout 5: Blank
    layout5 = SlideLayoutPart.new("Blank")
    layouts.append(layout5)

    # Layout 6: Title Only
    layout6 = SlideLayoutPart.new("Title Only")
    title_shape = Shape(
        id=2,
        name="Title",
        position=Position(x=457200, y=274638, cx=8229600, cy=1143000),
        placeholder=PlaceholderInfo(type="title"),
        text_frame=TextFrame(),
        preset_geometry="rect",
    )
    layout6.shape_tree.add_shape(title_shape)
    layouts.append(layout6)

    return layouts
