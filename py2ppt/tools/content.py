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
) -> None:
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

    Example:
        >>> add_text_box(pres, 1, "Note", "1in", "6in", "2in", "0.5in")
    """
    slide = presentation.get_slide(slide_number)
    slide.add_text_box(
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
