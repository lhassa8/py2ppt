"""Content manipulation tool functions.

Functions for setting text content in slides with support for rich text formatting.
"""

from __future__ import annotations

from typing import Any

from ..core.presentation import Presentation

# Type alias for rich text: either a string or list of formatted segments
RichText = str | list[dict[str, Any]]


def set_title(
    presentation: Presentation,
    slide_number: int,
    text: RichText,
    *,
    font_size: int | None = None,
    font_family: str | None = None,
    bold: bool = False,
    italic: bool = False,
    underline: bool = False,
    color: str | None = None,
) -> None:
    """Set the title of a slide.

    Supports both simple text and rich text with mixed formatting.

    Args:
        presentation: The presentation to modify
        slide_number: The slide number (1-indexed)
        text: The title text. Can be:
            - A simple string: "My Title"
            - A list of formatted segments for mixed formatting:
              [{"text": "Bold", "bold": True}, {"text": " Normal"}]
        font_size: Font size in points (e.g., 32). Applied to simple text.
        font_family: Font family name (e.g., "Arial"). Applied to simple text.
        bold: Whether to make the text bold. Applied to simple text.
        italic: Whether to make the text italic. Applied to simple text.
        underline: Whether to underline the text. Applied to simple text.
        color: Color as hex ("#FF0000"), rgb("rgb(255,0,0)"), or name ("red")

    Rich text segment options:
        - text: The text content (required)
        - bold: True/False
        - italic: True/False
        - underline: True/False
        - strikethrough: True/False
        - superscript: True/False
        - subscript: True/False
        - color: Hex color ("#FF0000") or theme color ("accent1")
        - highlight: Highlight color as hex
        - font_size: Size in points
        - font_family: Font name
        - hyperlink: URL for hyperlink

    Example:
        >>> # Simple text
        >>> set_title(pres, 1, "Q4 Business Review")

        >>> # With formatting
        >>> set_title(pres, 2, "Key Metrics", color="#0066CC", bold=True)

        >>> # Rich text with mixed formatting
        >>> set_title(pres, 3, [
        ...     {"text": "Important: ", "bold": True, "color": "#FF0000"},
        ...     {"text": "Q4 Results", "italic": True},
        ... ])
    """
    slide = presentation.get_slide(slide_number)
    slide.set_title(
        text,
        font_size=font_size,
        font_family=font_family,
        bold=bold,
        italic=italic,
        underline=underline,
        color=color,
    )


def set_subtitle(
    presentation: Presentation,
    slide_number: int,
    text: RichText,
    *,
    font_size: int | None = None,
    font_family: str | None = None,
    bold: bool = False,
    italic: bool = False,
    underline: bool = False,
    color: str | None = None,
) -> None:
    """Set the subtitle of a slide.

    Supports both simple text and rich text with mixed formatting.

    Args:
        presentation: The presentation to modify
        slide_number: The slide number (1-indexed)
        text: The subtitle text (string or list of formatted segments)
        font_size: Font size in points
        font_family: Font family name
        bold: Whether to make the text bold
        italic: Whether to make the text italic
        underline: Whether to underline the text
        color: Color as hex, rgb, or name

    Example:
        >>> set_subtitle(pres, 1, "Prepared by Analytics Team")

        >>> # Rich text
        >>> set_subtitle(pres, 1, [
        ...     {"text": "Prepared by "},
        ...     {"text": "Analytics Team", "bold": True},
        ... ])
    """
    slide = presentation.get_slide(slide_number)
    slide.set_subtitle(
        text,
        font_size=font_size,
        font_family=font_family,
        bold=bold,
        italic=italic,
        underline=underline,
        color=color,
    )


def set_body(
    presentation: Presentation,
    slide_number: int,
    content: str | list[str | RichText],
    *,
    levels: list[int] | None = None,
    font_size: int | None = None,
    font_family: str | None = None,
    color: str | None = None,
) -> None:
    """Set the body content of a slide.

    Supports both simple text and rich text with mixed formatting per bullet.

    Args:
        presentation: The presentation to modify
        slide_number: The slide number (1-indexed)
        content: Single string, list of bullet points, or list with rich text.
            Each item can be a string or list of formatted segments.
        levels: Optional list of indent levels (0-8) for each bullet.
                Default is 0 (top level) for all items.
        font_size: Font size in points (applied to simple text)
        font_family: Font family name (applied to simple text)
        color: Color as hex, rgb, or name (applied to simple text)

    Example:
        >>> # Simple bullets
        >>> set_body(pres, 2, [
        ...     "Revenue up 20%",
        ...     "New markets opened",
        ...     "Customer satisfaction at 95%"
        ... ])

        >>> # With nested bullets
        >>> set_body(pres, 3, [
        ...     "Main point",
        ...     "Sub-point 1",
        ...     "Sub-point 2",
        ...     "Another main point"
        ... ], levels=[0, 1, 1, 0])

        >>> # Rich text bullets
        >>> set_body(pres, 4, [
        ...     [{"text": "Important: ", "bold": True}, {"text": "Revenue up"}],
        ...     "Normal bullet point",
        ...     [{"text": "See ", "color": "#666666"},
        ...      {"text": "report", "hyperlink": "https://example.com"}],
        ... ])
    """
    slide = presentation.get_slide(slide_number)
    slide.set_body(
        content,
        levels=levels,
        font_size=font_size,
        font_family=font_family,
        color=color,
    )


def add_bullet(
    presentation: Presentation,
    slide_number: int,
    text: RichText,
    *,
    level: int = 0,
    font_size: int | None = None,
    font_family: str | None = None,
    bold: bool = False,
    italic: bool = False,
    color: str | None = None,
) -> None:
    """Add a bullet point to the slide body.

    Appends a new bullet to existing body content.
    Supports rich text with mixed formatting.

    Args:
        presentation: The presentation to modify
        slide_number: The slide number (1-indexed)
        text: The bullet text (string or list of formatted segments)
        level: Indent level (0-8). 0 is top level.
        font_size: Font size in points
        font_family: Font family name
        bold: Whether to make the text bold
        italic: Whether to make the text italic
        color: Color as hex, rgb, or name

    Example:
        >>> add_bullet(pres, 2, "Additional point")
        >>> add_bullet(pres, 2, "Sub-point", level=1)

        >>> # Rich text bullet
        >>> add_bullet(pres, 2, [
        ...     {"text": "Key: ", "bold": True},
        ...     {"text": "value description"},
        ... ])
    """
    slide = presentation.get_slide(slide_number)
    slide.add_bullet(
        text,
        level=level,
        font_size=font_size,
        font_family=font_family,
        bold=bold,
        italic=italic,
        color=color,
    )


def set_placeholder_text(
    presentation: Presentation,
    slide_number: int,
    placeholder: str,
    text: RichText,
    *,
    font_size: int | None = None,
    font_family: str | None = None,
    bold: bool = False,
    italic: bool = False,
    underline: bool = False,
    color: str | None = None,
) -> None:
    """Set text in a specific placeholder.

    Supports both simple text and rich text with mixed formatting.

    Args:
        presentation: The presentation to modify
        slide_number: The slide number (1-indexed)
        placeholder: Placeholder type or name. Common values:
                    "title", "subtitle", "body", "content",
                    "footer", "date", "slide_number"
                    For multiple placeholders of same type: "body_1", "body_2"
        text: The text content (string or list of formatted segments)
        font_size: Font size in points
        font_family: Font family name
        bold: Whether to make the text bold
        italic: Whether to make the text italic
        underline: Whether to underline the text
        color: Color as hex, rgb, or name

    Example:
        >>> set_placeholder_text(pres, 2, "body_1", "Left column content")
        >>> set_placeholder_text(pres, 2, "body_2", "Right column content")

        >>> # Rich text
        >>> set_placeholder_text(pres, 2, "footer", [
        ...     {"text": "Confidential", "color": "#FF0000", "bold": True},
        ... ])
    """
    slide = presentation.get_slide(slide_number)
    slide.set_placeholder_text(
        placeholder,
        text,
        font_size=font_size,
        font_family=font_family,
        bold=bold,
        italic=italic,
        underline=underline,
        color=color,
    )


def add_text_box(
    presentation: Presentation,
    slide_number: int,
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
) -> int:
    """Add a text box at a specific position.

    Args:
        presentation: The presentation to modify
        slide_number: The slide number (1-indexed)
        text: The text content
        left: Left position (e.g., "1in", "2.5cm", or EMU value)
        top: Top position
        width: Width
        height: Height
        font_size: Font size in points
        font_family: Font family name
        bold: Whether to make the text bold
        color: Color as hex, rgb, or name

    Returns:
        The shape ID of the created text box

    Example:
        >>> shape_id = add_text_box(pres, 1, "Note", "1in", "6in", "2in", "0.5in")
    """
    slide = presentation.get_slide(slide_number)
    shape = slide.add_text_box(
        text,
        left,
        top,
        width,
        height,
        font_size=font_size,
        font_family=font_family,
        bold=bold,
        color=color,
    )
    return shape.id


def set_notes(
    presentation: Presentation,
    slide_number: int,
    text: str,
) -> None:
    """Set speaker notes for a slide.

    Speaker notes appear in presenter view and are useful for
    adding talking points, reminders, or detailed information.

    Args:
        presentation: The presentation to modify
        slide_number: The slide number (1-indexed)
        text: The notes text. Use newlines for multiple paragraphs.

    Example:
        >>> set_notes(pres, 1, "Welcome the audience and introduce the topic.")

        >>> set_notes(pres, 2, '''Key talking points:
        ... - Emphasize the 20% growth
        ... - Mention new market expansion
        ... - Highlight customer feedback''')
    """
    from ..oxml.notes import create_notes_slide

    create_notes_slide(presentation._package, slide_number, text)
    presentation._dirty = True


def get_notes(
    presentation: Presentation,
    slide_number: int,
) -> str:
    """Get speaker notes from a slide.

    Args:
        presentation: The presentation to read from
        slide_number: The slide number (1-indexed)

    Returns:
        The notes text, or empty string if no notes exist.

    Example:
        >>> notes = get_notes(pres, 1)
        >>> if notes:
        ...     print(f"Notes: {notes}")
    """
    from ..oxml.notes import get_notes_slide

    notes_part = get_notes_slide(presentation._package, slide_number)
    if notes_part:
        return notes_part.get_text()
    return ""


def append_notes(
    presentation: Presentation,
    slide_number: int,
    text: str,
) -> None:
    """Append text to existing speaker notes.

    If the slide has no notes, creates new notes with the given text.

    Args:
        presentation: The presentation to modify
        slide_number: The slide number (1-indexed)
        text: Text to append to existing notes

    Example:
        >>> append_notes(pres, 1, "Additional point to mention.")
    """
    from ..oxml.notes import create_notes_slide, get_notes_slide

    notes_part = get_notes_slide(presentation._package, slide_number)
    if notes_part:
        notes_part.append_text(text)
        # Save the updated notes - recreate with full text
        full_text = notes_part.get_text()
        create_notes_slide(presentation._package, slide_number, full_text)
    else:
        create_notes_slide(presentation._package, slide_number, text)

    presentation._dirty = True


def find_text(
    presentation: Presentation,
    search_text: str,
    *,
    case_sensitive: bool = False,
    whole_word: bool = False,
) -> list[dict]:
    """Find text in the presentation.

    Searches through all slides and returns locations of matches.

    Args:
        presentation: The presentation to search
        search_text: Text to find
        case_sensitive: Match case exactly (default False)
        whole_word: Match whole words only (default False)

    Returns:
        List of dicts with match information:
        - slide: Slide number
        - shape: Shape name
        - text: The matched text with context

    Example:
        >>> matches = find_text(pres, "revenue")
        >>> for m in matches:
        ...     print(f"Slide {m['slide']}: {m['text']}")
    """
    import re

    results = []

    # Build regex pattern
    pattern = search_text
    if not case_sensitive:
        pattern = f"(?i){pattern}"
    if whole_word:
        pattern = rf"\b{pattern}\b"

    regex = re.compile(pattern)

    for slide_num in range(1, presentation.slide_count + 1):
        slide = presentation.get_slide(slide_num)

        for shape in slide.shapes:
            # Get text from shape
            shape_text = ""
            if hasattr(shape, "text_frame") and shape.text_frame:
                for para in shape.text_frame.paragraphs:
                    for run in para.runs:
                        if run.text:
                            shape_text += run.text

            if shape_text and regex.search(shape_text):
                # Find matches with context
                for match in regex.finditer(shape_text):
                    start = max(0, match.start() - 20)
                    end = min(len(shape_text), match.end() + 20)
                    context = shape_text[start:end]
                    if start > 0:
                        context = "..." + context
                    if end < len(shape_text):
                        context = context + "..."

                    results.append({
                        "slide": slide_num,
                        "shape": shape.name,
                        "text": context,
                        "match": match.group(),
                    })

    return results


def replace_text(
    presentation: Presentation,
    old_text: str,
    new_text: str,
    *,
    case_sensitive: bool = False,
    whole_word: bool = False,
) -> int:
    """Replace text throughout the presentation.

    Args:
        presentation: The presentation to modify
        old_text: Text to find and replace
        new_text: Replacement text
        case_sensitive: Match case exactly (default False)
        whole_word: Match whole words only (default False)

    Returns:
        Number of replacements made

    Example:
        >>> count = replace_text(pres, "2023", "2024")
        >>> print(f"Replaced {count} occurrences")
    """
    import re

    # Build regex pattern
    pattern = re.escape(old_text)
    if not case_sensitive:
        pattern = f"(?i){pattern}"
    if whole_word:
        pattern = rf"\b{pattern}\b"

    regex = re.compile(pattern)
    total_replacements = 0

    for slide_num in range(1, presentation.slide_count + 1):
        slide = presentation.get_slide(slide_num)
        slide_modified = False

        for shape in slide.shapes:
            if hasattr(shape, "text_frame") and shape.text_frame:
                for para in shape.text_frame.paragraphs:
                    for run in para.runs:
                        if run.text and regex.search(run.text):
                            new_run_text, count = regex.subn(new_text, run.text)
                            if count > 0:
                                run.text = new_run_text
                                total_replacements += count
                                slide_modified = True

        if slide_modified:
            slide._save()

    return total_replacements


def replace_all(
    presentation: Presentation,
    replacements: dict[str, str],
    *,
    case_sensitive: bool = False,
) -> dict[str, int]:
    """Replace multiple text strings at once.

    More efficient than calling replace_text multiple times.

    Args:
        presentation: The presentation to modify
        replacements: Dict mapping old text to new text
        case_sensitive: Match case exactly (default False)

    Returns:
        Dict mapping each old text to number of replacements made

    Example:
        >>> counts = replace_all(pres, {
        ...     "Company Name": "Acme Corp",
        ...     "2023": "2024",
        ...     "Q3": "Q4",
        ... })
        >>> for old, count in counts.items():
        ...     print(f"Replaced '{old}' {count} times")
    """
    results = {}
    for old_text, new_text in replacements.items():
        count = replace_text(
            presentation, old_text, new_text,
            case_sensitive=case_sensitive
        )
        results[old_text] = count
    return results


# ============================================================================
# Headers and Footers
# ============================================================================


def set_footer(
    presentation: Presentation,
    text: str,
    *,
    apply_to_all: bool = True,
    slide_number: int | None = None,
) -> None:
    """Set the footer text for slides.

    Args:
        presentation: The presentation to modify
        text: Footer text to display
        apply_to_all: Apply to all slides (default True)
        slide_number: Specific slide to apply to (if apply_to_all is False)

    Example:
        >>> set_footer(pres, "Confidential - Internal Use Only")
        >>> set_footer(pres, "Draft", apply_to_all=False, slide_number=1)
    """
    from lxml import etree

    from ..oxml.ns import CONTENT_TYPE, qn

    if apply_to_all:
        # Update presentation-level header/footer settings
        pres_part = presentation._presentation
        pres_elem = pres_part._element

        # Find or create hf element
        hf = pres_elem.find(qn("p:hf"))
        if hf is None:
            # Insert after custShowLst or sldIdLst if they exist
            insert_after = pres_elem.find(qn("p:custShowLst"))
            if insert_after is None:
                insert_after = pres_elem.find(qn("p:sldIdLst"))

            hf = etree.Element(qn("p:hf"))
            if insert_after is not None:
                idx = list(pres_elem).index(insert_after) + 1
                pres_elem.insert(idx, hf)
            else:
                pres_elem.append(hf)

        hf.set("ftr", "1")  # Enable footer

        # Save presentation
        presentation._package.set_part(
            "ppt/presentation.xml",
            pres_part.to_xml(),
            CONTENT_TYPE.PRESENTATION,
        )

        # Update footer placeholder text in each slide
        for i in range(1, presentation.slide_count + 1):
            _set_slide_footer(presentation, i, text)
    else:
        if slide_number is None:
            raise ValueError("slide_number required when apply_to_all is False")
        _set_slide_footer(presentation, slide_number, text)

    presentation._dirty = True


def _set_slide_footer(presentation: Presentation, slide_number: int, text: str) -> None:
    """Set footer text on a specific slide."""
    from lxml import etree

    from ..oxml.ns import qn
    from ..oxml.slide import update_slide_in_package

    slide = presentation.get_slide(slide_number)
    slide_part = slide._part
    slide_elem = slide_part._element

    # Find shape tree
    sp_tree = slide_elem.find(f".//{qn('p:spTree')}")
    if sp_tree is None:
        return

    # Look for footer placeholder
    footer_found = False
    for sp in sp_tree.findall(qn("p:sp")):
        ph = sp.find(f".//{qn('p:ph')}")
        if ph is not None and ph.get("type") == "ftr":
            # Found footer placeholder - set text
            tx_body = sp.find(qn("p:txBody"))
            if tx_body is None:
                tx_body = etree.SubElement(sp, qn("p:txBody"))
                etree.SubElement(tx_body, qn("a:bodyPr"))
                etree.SubElement(tx_body, qn("a:lstStyle"))

            # Remove existing paragraphs
            for p in tx_body.findall(qn("a:p")):
                tx_body.remove(p)

            # Add new paragraph with text
            p = etree.SubElement(tx_body, qn("a:p"))
            r = etree.SubElement(p, qn("a:r"))
            t = etree.SubElement(r, qn("a:t"))
            t.text = text

            footer_found = True
            break

    # If no footer placeholder exists, we could add one (optional)
    if not footer_found:
        # Footer placeholder needs to be defined in layout/master
        # For now, just return without error
        pass

    update_slide_in_package(
        presentation._package,
        slide_number,
        slide_part,
    )


def set_slide_number_visibility(
    presentation: Presentation,
    visible: bool = True,
    *,
    start_from: int = 1,
) -> None:
    """Control whether slide numbers are displayed.

    Args:
        presentation: The presentation to modify
        visible: Whether slide numbers should be visible
        start_from: Starting number for slides (default 1)

    Example:
        >>> set_slide_number_visibility(pres, visible=True)
        >>> set_slide_number_visibility(pres, visible=True, start_from=0)
    """
    from lxml import etree

    from ..oxml.ns import CONTENT_TYPE, qn

    pres_part = presentation._presentation
    pres_elem = pres_part._element

    # Find or create hf element
    hf = pres_elem.find(qn("p:hf"))
    if hf is None:
        hf = etree.Element(qn("p:hf"))
        # Insert at appropriate position
        sld_id_lst = pres_elem.find(qn("p:sldIdLst"))
        if sld_id_lst is not None:
            idx = list(pres_elem).index(sld_id_lst) + 1
            pres_elem.insert(idx, hf)
        else:
            pres_elem.append(hf)

    hf.set("sldNum", "1" if visible else "0")

    # Set first slide number
    if start_from != 1:
        pres_elem.set("firstSlideNum", str(start_from))

    presentation._package.set_part(
        "ppt/presentation.xml",
        pres_part.to_xml(),
        CONTENT_TYPE.PRESENTATION,
    )
    presentation._dirty = True


def set_date_visibility(
    presentation: Presentation,
    visible: bool = True,
    *,
    auto_update: bool = True,
    date_format: str | None = None,
    fixed_date: str | None = None,
) -> None:
    """Control whether date is displayed in slides.

    Args:
        presentation: The presentation to modify
        visible: Whether date should be visible
        auto_update: Update date automatically (default True)
        date_format: Date format string (e.g., "MMMM d, yyyy")
        fixed_date: Fixed date text (used when auto_update is False)

    Example:
        >>> set_date_visibility(pres, visible=True)
        >>> set_date_visibility(pres, visible=True, auto_update=False, fixed_date="Q4 2024")
    """
    from lxml import etree

    from ..oxml.ns import CONTENT_TYPE, qn

    pres_part = presentation._presentation
    pres_elem = pres_part._element

    # Find or create hf element
    hf = pres_elem.find(qn("p:hf"))
    if hf is None:
        hf = etree.Element(qn("p:hf"))
        sld_id_lst = pres_elem.find(qn("p:sldIdLst"))
        if sld_id_lst is not None:
            idx = list(pres_elem).index(sld_id_lst) + 1
            pres_elem.insert(idx, hf)
        else:
            pres_elem.append(hf)

    hf.set("dt", "1" if visible else "0")

    presentation._package.set_part(
        "ppt/presentation.xml",
        pres_part.to_xml(),
        CONTENT_TYPE.PRESENTATION,
    )
    presentation._dirty = True


def get_header_footer_settings(presentation: Presentation) -> dict:
    """Get current header/footer settings.

    Returns:
        Dict with settings:
        - footer_visible: bool
        - slide_number_visible: bool
        - date_visible: bool
        - first_slide_number: int

    Example:
        >>> settings = get_header_footer_settings(pres)
        >>> if settings["slide_number_visible"]:
        ...     print("Slide numbers are on")
    """
    from ..oxml.ns import qn

    pres_part = presentation._presentation
    pres_elem = pres_part._element

    hf = pres_elem.find(qn("p:hf"))

    settings = {
        "footer_visible": False,
        "slide_number_visible": False,
        "date_visible": False,
        "first_slide_number": 1,
    }

    if hf is not None:
        settings["footer_visible"] = hf.get("ftr") == "1"
        settings["slide_number_visible"] = hf.get("sldNum") == "1"
        settings["date_visible"] = hf.get("dt") == "1"

    first_num = pres_elem.get("firstSlideNum")
    if first_num:
        settings["first_slide_number"] = int(first_num)

    return settings


# ============================================================================
# Hyperlinks
# ============================================================================


def add_hyperlink(
    presentation: Presentation,
    slide_number: int,
    shape_id: int,
    *,
    url: str | None = None,
    slide: int | None = None,
    action: str | None = None,
    tooltip: str | None = None,
) -> bool:
    """Add a hyperlink to a shape.

    Links can be external URLs, internal slide references, or navigation actions.

    Args:
        presentation: The presentation to modify
        slide_number: The slide number (1-indexed)
        shape_id: The shape ID to add the link to
        url: External URL (e.g., "https://example.com")
        slide: Internal slide number to link to (1-indexed)
        action: Navigation action. One of:
            - "next_slide": Go to next slide
            - "previous_slide": Go to previous slide
            - "first_slide": Go to first slide
            - "last_slide": Go to last slide
            - "end_show": End the slideshow
        tooltip: Tooltip text shown on hover

    Returns:
        True if successful, False if shape not found

    Note:
        Only one of url, slide, or action should be specified.

    Example:
        >>> # External link
        >>> add_hyperlink(pres, 1, shape_id, url="https://example.com")

        >>> # Internal slide link
        >>> add_hyperlink(pres, 1, shape_id, slide=5, tooltip="Go to Summary")

        >>> # Navigation action
        >>> add_hyperlink(pres, 1, shape_id, action="next_slide")
    """
    from lxml import etree

    from ..oxml.ns import CONTENT_TYPE, REL_TYPE, qn

    pkg = presentation._package

    # Get slide element
    slide_refs = presentation._presentation.get_slide_refs()
    if slide_number < 1 or slide_number > len(slide_refs):
        return False

    slide_ref = slide_refs[slide_number - 1]
    pres_rels = pkg.get_part_rels("ppt/presentation.xml")
    rel = pres_rels.get(slide_ref.r_id)

    if rel.target.startswith("/"):
        slide_path = rel.target.lstrip("/")
    else:
        slide_path = f"ppt/{rel.target}"

    slide_xml = pkg.get_part(slide_path)
    if slide_xml is None:
        return False

    slide_elem = etree.fromstring(slide_xml)

    # Find shape by ID
    shape_elem = None
    sp_tree = slide_elem.find(f".//{qn('p:spTree')}")
    if sp_tree is None:
        return False

    # Check shapes (p:sp)
    for sp in sp_tree.findall(f".//{qn('p:sp')}"):
        c_nv_pr = sp.find(f".//{qn('p:cNvPr')}")
        if c_nv_pr is not None and c_nv_pr.get("id") == str(shape_id):
            shape_elem = sp
            break

    # Check pictures (p:pic)
    if shape_elem is None:
        for pic in sp_tree.findall(f".//{qn('p:pic')}"):
            c_nv_pr = pic.find(f".//{qn('p:cNvPr')}")
            if c_nv_pr is not None and c_nv_pr.get("id") == str(shape_id):
                shape_elem = pic
                break

    if shape_elem is None:
        return False

    # Find cNvPr element where hyperlink goes
    c_nv_pr = shape_elem.find(f".//{qn('p:cNvPr')}")
    if c_nv_pr is None:
        return False

    # Remove existing hyperlink if any
    existing_hlink = c_nv_pr.find(qn("a:hlinkClick"))
    if existing_hlink is not None:
        c_nv_pr.remove(existing_hlink)

    # Create hyperlink element
    hlink_click = etree.SubElement(c_nv_pr, qn("a:hlinkClick"))

    slide_rels = pkg.get_part_rels(slide_path)

    if url:
        # External URL
        r_id = slide_rels.add(
            rel_type=REL_TYPE.HYPERLINK,
            target=url,
            target_mode="External",
        )
        hlink_click.set(qn("r:id"), r_id)

    elif slide is not None:
        # Internal slide link
        # PowerPoint uses action with slide reference
        hlink_click.set("action", "ppaction://hlinksldjump")
        # Add relationship to target slide
        target_slide_path = f"slide{slide}.xml"
        r_id = slide_rels.add(
            rel_type=REL_TYPE.SLIDE,
            target=target_slide_path,
        )
        hlink_click.set(qn("r:id"), r_id)

    elif action:
        # Navigation action
        action_map = {
            "next_slide": "ppaction://hlinkshowjump?jump=nextslide",
            "previous_slide": "ppaction://hlinkshowjump?jump=previousslide",
            "first_slide": "ppaction://hlinkshowjump?jump=firstslide",
            "last_slide": "ppaction://hlinkshowjump?jump=lastslide",
            "end_show": "ppaction://hlinkshowjump?jump=endshow",
        }
        action_uri = action_map.get(action.lower().replace("-", "_").replace(" ", "_"))
        if action_uri:
            hlink_click.set("action", action_uri)
        else:
            return False

    if tooltip:
        hlink_click.set("tooltip", tooltip)

    # Save changes
    pkg.set_part_rels(slide_path, slide_rels)

    xml_bytes = etree.tostring(
        slide_elem,
        xml_declaration=True,
        encoding="UTF-8",
        standalone=True,
    )
    pkg.set_part(slide_path, xml_bytes, CONTENT_TYPE.SLIDE)
    presentation._dirty = True

    return True


def remove_hyperlink(
    presentation: Presentation,
    slide_number: int,
    shape_id: int,
) -> bool:
    """Remove a hyperlink from a shape.

    Args:
        presentation: The presentation to modify
        slide_number: The slide number (1-indexed)
        shape_id: The shape ID to remove the link from

    Returns:
        True if hyperlink was removed, False if not found

    Example:
        >>> remove_hyperlink(pres, 1, shape_id)
    """
    from lxml import etree

    from ..oxml.ns import CONTENT_TYPE, qn

    pkg = presentation._package

    # Get slide element
    slide_refs = presentation._presentation.get_slide_refs()
    if slide_number < 1 or slide_number > len(slide_refs):
        return False

    slide_ref = slide_refs[slide_number - 1]
    pres_rels = pkg.get_part_rels("ppt/presentation.xml")
    rel = pres_rels.get(slide_ref.r_id)

    if rel.target.startswith("/"):
        slide_path = rel.target.lstrip("/")
    else:
        slide_path = f"ppt/{rel.target}"

    slide_xml = pkg.get_part(slide_path)
    if slide_xml is None:
        return False

    slide_elem = etree.fromstring(slide_xml)

    # Find shape by ID
    shape_elem = None
    sp_tree = slide_elem.find(f".//{qn('p:spTree')}")
    if sp_tree is None:
        return False

    for sp in sp_tree.findall(f".//{qn('p:sp')}"):
        c_nv_pr = sp.find(f".//{qn('p:cNvPr')}")
        if c_nv_pr is not None and c_nv_pr.get("id") == str(shape_id):
            shape_elem = sp
            break

    if shape_elem is None:
        for pic in sp_tree.findall(f".//{qn('p:pic')}"):
            c_nv_pr = pic.find(f".//{qn('p:cNvPr')}")
            if c_nv_pr is not None and c_nv_pr.get("id") == str(shape_id):
                shape_elem = pic
                break

    if shape_elem is None:
        return False

    # Find and remove hyperlink
    c_nv_pr = shape_elem.find(f".//{qn('p:cNvPr')}")
    if c_nv_pr is None:
        return False

    hlink = c_nv_pr.find(qn("a:hlinkClick"))
    if hlink is None:
        return False

    c_nv_pr.remove(hlink)

    # Save changes
    xml_bytes = etree.tostring(
        slide_elem,
        xml_declaration=True,
        encoding="UTF-8",
        standalone=True,
    )
    pkg.set_part(slide_path, xml_bytes, CONTENT_TYPE.SLIDE)
    presentation._dirty = True

    return True


def get_hyperlinks(
    presentation: Presentation,
    slide_number: int,
) -> list[dict]:
    """Get all hyperlinks on a slide.

    Args:
        presentation: The presentation to inspect
        slide_number: The slide number (1-indexed)

    Returns:
        List of dicts with hyperlink info:
        - shape_id: ID of the shape with the link
        - shape_name: Name of the shape
        - url: External URL (if applicable)
        - slide: Target slide number (if internal link)
        - action: Navigation action (if applicable)
        - tooltip: Tooltip text (if set)

    Example:
        >>> links = get_hyperlinks(pres, 1)
        >>> for link in links:
        ...     print(f"Shape {link['shape_id']}: {link.get('url', link.get('action'))}")
    """
    from lxml import etree

    from ..oxml.ns import REL_TYPE, qn

    pkg = presentation._package
    links = []

    # Get slide element
    slide_refs = presentation._presentation.get_slide_refs()
    if slide_number < 1 or slide_number > len(slide_refs):
        return links

    slide_ref = slide_refs[slide_number - 1]
    pres_rels = pkg.get_part_rels("ppt/presentation.xml")
    rel = pres_rels.get(slide_ref.r_id)

    if rel.target.startswith("/"):
        slide_path = rel.target.lstrip("/")
    else:
        slide_path = f"ppt/{rel.target}"

    slide_xml = pkg.get_part(slide_path)
    if slide_xml is None:
        return links

    slide_elem = etree.fromstring(slide_xml)
    slide_rels = pkg.get_part_rels(slide_path)

    # Find all shapes with hyperlinks
    sp_tree = slide_elem.find(f".//{qn('p:spTree')}")
    if sp_tree is None:
        return links

    # Check shapes and pictures
    for shape_type in ["p:sp", "p:pic"]:
        for shape in sp_tree.findall(f".//{qn(shape_type)}"):
            c_nv_pr = shape.find(f".//{qn('p:cNvPr')}")
            if c_nv_pr is None:
                continue

            hlink = c_nv_pr.find(qn("a:hlinkClick"))
            if hlink is None:
                continue

            shape_id = int(c_nv_pr.get("id", "0"))
            shape_name = c_nv_pr.get("name", "")

            link_info = {
                "shape_id": shape_id,
                "shape_name": shape_name,
            }

            # Get tooltip
            tooltip = hlink.get("tooltip")
            if tooltip:
                link_info["tooltip"] = tooltip

            # Get relationship ID
            r_id = hlink.get(qn("r:id"))
            action = hlink.get("action")

            if r_id:
                # Look up the relationship
                link_rel = slide_rels.get(r_id)
                if link_rel:
                    if link_rel.rel_type == REL_TYPE.HYPERLINK:
                        link_info["url"] = link_rel.target
                    elif link_rel.rel_type == REL_TYPE.SLIDE:
                        # Parse slide number from target
                        target = link_rel.target
                        if "slide" in target.lower():
                            import re
                            match = re.search(r"slide(\d+)", target, re.IGNORECASE)
                            if match:
                                link_info["slide"] = int(match.group(1))

            if action:
                # Parse action
                action_reverse = {
                    "ppaction://hlinkshowjump?jump=nextslide": "next_slide",
                    "ppaction://hlinkshowjump?jump=previousslide": "previous_slide",
                    "ppaction://hlinkshowjump?jump=firstslide": "first_slide",
                    "ppaction://hlinkshowjump?jump=lastslide": "last_slide",
                    "ppaction://hlinkshowjump?jump=endshow": "end_show",
                    "ppaction://hlinksldjump": "slide_jump",
                }
                parsed_action = action_reverse.get(action.lower())
                if parsed_action and parsed_action != "slide_jump":
                    link_info["action"] = parsed_action

            links.append(link_info)

    return links


# ============================================================================
# Text Columns
# ============================================================================


def set_text_columns(
    presentation: Presentation,
    slide_number: int,
    shape_id: int,
    num_columns: int = 2,
    *,
    spacing: str | int = "0.5in",
) -> bool:
    """Set the number of text columns in a shape.

    Args:
        presentation: The presentation to modify
        slide_number: The slide number (1-indexed)
        shape_id: The shape ID to modify
        num_columns: Number of columns (1-10)
        spacing: Space between columns (e.g., "0.5in", "1cm")

    Returns:
        True if successful, False if shape not found

    Example:
        >>> # Two-column text box
        >>> set_text_columns(pres, 1, shape_id, num_columns=2)

        >>> # Three columns with custom spacing
        >>> set_text_columns(pres, 1, shape_id, num_columns=3, spacing="0.25in")
    """
    from lxml import etree

    from ..oxml.ns import CONTENT_TYPE, qn
    from ..utils.units import parse_length

    if num_columns < 1 or num_columns > 10:
        raise ValueError("num_columns must be between 1 and 10")

    pkg = presentation._package

    # Get slide element
    slide_refs = presentation._presentation.get_slide_refs()
    if slide_number < 1 or slide_number > len(slide_refs):
        return False

    slide_ref = slide_refs[slide_number - 1]
    pres_rels = pkg.get_part_rels("ppt/presentation.xml")
    rel = pres_rels.get(slide_ref.r_id)

    if rel.target.startswith("/"):
        slide_path = rel.target.lstrip("/")
    else:
        slide_path = f"ppt/{rel.target}"

    slide_xml = pkg.get_part(slide_path)
    if slide_xml is None:
        return False

    slide_elem = etree.fromstring(slide_xml)

    # Find shape by ID
    sp_tree = slide_elem.find(f".//{qn('p:spTree')}")
    if sp_tree is None:
        return False

    shape_elem = None
    for sp in sp_tree.findall(f".//{qn('p:sp')}"):
        c_nv_pr = sp.find(f".//{qn('p:cNvPr')}")
        if c_nv_pr is not None and c_nv_pr.get("id") == str(shape_id):
            shape_elem = sp
            break

    if shape_elem is None:
        return False

    # Find or create txBody/bodyPr
    tx_body = shape_elem.find(qn("p:txBody"))
    if tx_body is None:
        return False

    body_pr = tx_body.find(qn("a:bodyPr"))
    if body_pr is None:
        body_pr = etree.Element(qn("a:bodyPr"))
        tx_body.insert(0, body_pr)

    # Set columns
    body_pr.set("numCol", str(num_columns))

    # Set spacing (in EMUs)
    spacing_emu = int(parse_length(spacing)) if isinstance(spacing, str) else spacing
    body_pr.set("spcCol", str(spacing_emu))

    # Save
    xml_bytes = etree.tostring(
        slide_elem,
        xml_declaration=True,
        encoding="UTF-8",
        standalone=True,
    )
    pkg.set_part(slide_path, xml_bytes, CONTENT_TYPE.SLIDE)
    presentation._dirty = True

    return True


def get_text_columns(
    presentation: Presentation,
    slide_number: int,
    shape_id: int,
) -> dict | None:
    """Get text column settings for a shape.

    Args:
        presentation: The presentation to inspect
        slide_number: The slide number (1-indexed)
        shape_id: The shape ID to inspect

    Returns:
        Dict with column settings or None if shape not found:
        - columns: Number of columns
        - spacing: Space between columns (EMUs)

    Example:
        >>> cols = get_text_columns(pres, 1, shape_id)
        >>> print(f"Columns: {cols['columns']}")
    """
    from lxml import etree

    from ..oxml.ns import qn

    pkg = presentation._package

    # Get slide element
    slide_refs = presentation._presentation.get_slide_refs()
    if slide_number < 1 or slide_number > len(slide_refs):
        return None

    slide_ref = slide_refs[slide_number - 1]
    pres_rels = pkg.get_part_rels("ppt/presentation.xml")
    rel = pres_rels.get(slide_ref.r_id)

    if rel.target.startswith("/"):
        slide_path = rel.target.lstrip("/")
    else:
        slide_path = f"ppt/{rel.target}"

    slide_xml = pkg.get_part(slide_path)
    if slide_xml is None:
        return None

    slide_elem = etree.fromstring(slide_xml)

    # Find shape by ID
    sp_tree = slide_elem.find(f".//{qn('p:spTree')}")
    if sp_tree is None:
        return None

    for sp in sp_tree.findall(f".//{qn('p:sp')}"):
        c_nv_pr = sp.find(f".//{qn('p:cNvPr')}")
        if c_nv_pr is not None and c_nv_pr.get("id") == str(shape_id):
            tx_body = sp.find(qn("p:txBody"))
            if tx_body is None:
                return {"columns": 1, "spacing": 0}

            body_pr = tx_body.find(qn("a:bodyPr"))
            if body_pr is None:
                return {"columns": 1, "spacing": 0}

            return {
                "columns": int(body_pr.get("numCol", "1")),
                "spacing": int(body_pr.get("spcCol", "0")),
            }

    return None


# ============================================================================
# Shape Text
# ============================================================================


def set_shape_text(
    presentation: Presentation,
    slide_number: int,
    shape_id: int,
    text: str,
    *,
    font_size: int | None = None,
    bold: bool = False,
    italic: bool = False,
    color: str | None = None,
    align: str = "center",
    vertical_align: str = "middle",
) -> bool:
    """Set or replace text in a shape.

    Works with auto-shapes (rectangles, ellipses, etc.) to add text.

    Args:
        presentation: The presentation to modify
        slide_number: The slide number (1-indexed)
        shape_id: The shape ID to modify
        text: Text to set
        font_size: Font size in points
        bold: Make text bold
        italic: Make text italic
        color: Text color as hex (e.g., "#FF0000")
        align: Horizontal alignment ("left", "center", "right")
        vertical_align: Vertical alignment ("top", "middle", "bottom")

    Returns:
        True if successful, False if shape not found

    Example:
        >>> # Add centered text to a shape
        >>> set_shape_text(pres, 1, shape_id, "Click Here",
        ...     font_size=24, bold=True, color="#FFFFFF")
    """
    from lxml import etree

    from ..oxml.ns import CONTENT_TYPE, qn

    pkg = presentation._package

    # Get slide element
    slide_refs = presentation._presentation.get_slide_refs()
    if slide_number < 1 or slide_number > len(slide_refs):
        return False

    slide_ref = slide_refs[slide_number - 1]
    pres_rels = pkg.get_part_rels("ppt/presentation.xml")
    rel = pres_rels.get(slide_ref.r_id)

    if rel.target.startswith("/"):
        slide_path = rel.target.lstrip("/")
    else:
        slide_path = f"ppt/{rel.target}"

    slide_xml = pkg.get_part(slide_path)
    if slide_xml is None:
        return False

    slide_elem = etree.fromstring(slide_xml)

    # Find shape by ID
    sp_tree = slide_elem.find(f".//{qn('p:spTree')}")
    if sp_tree is None:
        return False

    shape_elem = None
    for sp in sp_tree.findall(f".//{qn('p:sp')}"):
        c_nv_pr = sp.find(f".//{qn('p:cNvPr')}")
        if c_nv_pr is not None and c_nv_pr.get("id") == str(shape_id):
            shape_elem = sp
            break

    if shape_elem is None:
        return False

    # Find or create txBody
    tx_body = shape_elem.find(qn("p:txBody"))
    if tx_body is None:
        tx_body = etree.SubElement(shape_elem, qn("p:txBody"))

    # Clear existing content
    for child in list(tx_body):
        tx_body.remove(child)

    # Body properties
    body_pr = etree.SubElement(tx_body, qn("a:bodyPr"))

    # Vertical alignment
    v_align_map = {"top": "t", "middle": "ctr", "bottom": "b"}
    body_pr.set("anchor", v_align_map.get(vertical_align.lower(), "ctr"))

    # List style
    etree.SubElement(tx_body, qn("a:lstStyle"))

    # Paragraph
    p = etree.SubElement(tx_body, qn("a:p"))

    # Paragraph properties (alignment)
    p_pr = etree.SubElement(p, qn("a:pPr"))
    h_align_map = {"left": "l", "center": "ctr", "right": "r"}
    p_pr.set("algn", h_align_map.get(align.lower(), "ctr"))

    # Run
    r = etree.SubElement(p, qn("a:r"))

    # Run properties
    r_pr = etree.SubElement(r, qn("a:rPr"))
    r_pr.set("lang", "en-US")

    if font_size:
        r_pr.set("sz", str(font_size * 100))  # Size in centipoints
    if bold:
        r_pr.set("b", "1")
    if italic:
        r_pr.set("i", "1")

    if color:
        solid_fill = etree.SubElement(r_pr, qn("a:solidFill"))
        srgb_clr = etree.SubElement(solid_fill, qn("a:srgbClr"))
        srgb_clr.set("val", color.lstrip("#").upper())

    # Text
    t = etree.SubElement(r, qn("a:t"))
    t.text = text

    # End paragraph run properties
    etree.SubElement(p, qn("a:endParaRPr"))

    # Save
    xml_bytes = etree.tostring(
        slide_elem,
        xml_declaration=True,
        encoding="UTF-8",
        standalone=True,
    )
    pkg.set_part(slide_path, xml_bytes, CONTENT_TYPE.SLIDE)
    presentation._dirty = True

    return True


def get_shape_text(
    presentation: Presentation,
    slide_number: int,
    shape_id: int,
) -> str | None:
    """Get text from a shape.

    Args:
        presentation: The presentation to inspect
        slide_number: The slide number (1-indexed)
        shape_id: The shape ID to inspect

    Returns:
        Text content or None if shape not found or has no text

    Example:
        >>> text = get_shape_text(pres, 1, shape_id)
        >>> if text:
        ...     print(f"Shape text: {text}")
    """
    from lxml import etree

    from ..oxml.ns import qn

    pkg = presentation._package

    # Get slide element
    slide_refs = presentation._presentation.get_slide_refs()
    if slide_number < 1 or slide_number > len(slide_refs):
        return None

    slide_ref = slide_refs[slide_number - 1]
    pres_rels = pkg.get_part_rels("ppt/presentation.xml")
    rel = pres_rels.get(slide_ref.r_id)

    if rel.target.startswith("/"):
        slide_path = rel.target.lstrip("/")
    else:
        slide_path = f"ppt/{rel.target}"

    slide_xml = pkg.get_part(slide_path)
    if slide_xml is None:
        return None

    slide_elem = etree.fromstring(slide_xml)

    # Find shape by ID
    sp_tree = slide_elem.find(f".//{qn('p:spTree')}")
    if sp_tree is None:
        return None

    for sp in sp_tree.findall(f".//{qn('p:sp')}"):
        c_nv_pr = sp.find(f".//{qn('p:cNvPr')}")
        if c_nv_pr is not None and c_nv_pr.get("id") == str(shape_id):
            # Get all text elements
            text_parts = []
            for t in sp.findall(f".//{qn('a:t')}"):
                if t.text:
                    text_parts.append(t.text)
            return "".join(text_parts) if text_parts else None

    return None


def append_shape_text(
    presentation: Presentation,
    slide_number: int,
    shape_id: int,
    text: str,
    *,
    new_paragraph: bool = False,
    font_size: int | None = None,
    bold: bool = False,
    italic: bool = False,
    color: str | None = None,
) -> bool:
    """Append text to an existing shape.

    Args:
        presentation: The presentation to modify
        slide_number: The slide number (1-indexed)
        shape_id: The shape ID to modify
        text: Text to append
        new_paragraph: If True, start a new paragraph
        font_size: Font size in points
        bold: Make appended text bold
        italic: Make appended text italic
        color: Text color as hex

    Returns:
        True if successful, False if shape not found

    Example:
        >>> # Append to existing text
        >>> append_shape_text(pres, 1, shape_id, " - continued")

        >>> # Add new paragraph
        >>> append_shape_text(pres, 1, shape_id, "Next point", new_paragraph=True)
    """
    from lxml import etree

    from ..oxml.ns import CONTENT_TYPE, qn

    pkg = presentation._package

    # Get slide element
    slide_refs = presentation._presentation.get_slide_refs()
    if slide_number < 1 or slide_number > len(slide_refs):
        return False

    slide_ref = slide_refs[slide_number - 1]
    pres_rels = pkg.get_part_rels("ppt/presentation.xml")
    rel = pres_rels.get(slide_ref.r_id)

    if rel.target.startswith("/"):
        slide_path = rel.target.lstrip("/")
    else:
        slide_path = f"ppt/{rel.target}"

    slide_xml = pkg.get_part(slide_path)
    if slide_xml is None:
        return False

    slide_elem = etree.fromstring(slide_xml)

    # Find shape by ID
    sp_tree = slide_elem.find(f".//{qn('p:spTree')}")
    if sp_tree is None:
        return False

    shape_elem = None
    for sp in sp_tree.findall(f".//{qn('p:sp')}"):
        c_nv_pr = sp.find(f".//{qn('p:cNvPr')}")
        if c_nv_pr is not None and c_nv_pr.get("id") == str(shape_id):
            shape_elem = sp
            break

    if shape_elem is None:
        return False

    # Find txBody
    tx_body = shape_elem.find(qn("p:txBody"))
    if tx_body is None:
        # No text body, create one with set_shape_text
        return set_shape_text(
            presentation, slide_number, shape_id, text,
            font_size=font_size, bold=bold, italic=italic, color=color
        )

    if new_paragraph:
        # Add new paragraph
        p = etree.SubElement(tx_body, qn("a:p"))
    else:
        # Find last paragraph
        paragraphs = tx_body.findall(qn("a:p"))
        if paragraphs:
            p = paragraphs[-1]
            # Remove endParaRPr if exists (we'll add it back)
            end_pr = p.find(qn("a:endParaRPr"))
            if end_pr is not None:
                p.remove(end_pr)
        else:
            p = etree.SubElement(tx_body, qn("a:p"))

    # Add run
    r = etree.SubElement(p, qn("a:r"))

    # Run properties
    r_pr = etree.SubElement(r, qn("a:rPr"))
    r_pr.set("lang", "en-US")

    if font_size:
        r_pr.set("sz", str(font_size * 100))
    if bold:
        r_pr.set("b", "1")
    if italic:
        r_pr.set("i", "1")

    if color:
        solid_fill = etree.SubElement(r_pr, qn("a:solidFill"))
        srgb_clr = etree.SubElement(solid_fill, qn("a:srgbClr"))
        srgb_clr.set("val", color.lstrip("#").upper())

    # Text
    t = etree.SubElement(r, qn("a:t"))
    t.text = text

    # Add endParaRPr
    etree.SubElement(p, qn("a:endParaRPr"))

    # Save
    xml_bytes = etree.tostring(
        slide_elem,
        xml_declaration=True,
        encoding="UTF-8",
        standalone=True,
    )
    pkg.set_part(slide_path, xml_bytes, CONTENT_TYPE.SLIDE)
    presentation._dirty = True

    return True
