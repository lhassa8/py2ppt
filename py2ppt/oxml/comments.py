"""Comment handling for PowerPoint presentations.

PowerPoint stores comments in separate parts:
- ppt/comments/comment[N].xml - Comment content for each slide
- ppt/commentAuthors.xml - List of comment authors
"""

from dataclasses import dataclass, field
from datetime import datetime
from typing import TYPE_CHECKING

from lxml import etree

from .ns import NAMESPACES, qn

if TYPE_CHECKING:
    pass

# Content type for comments
COMMENT_CONTENT_TYPE = (
    "application/vnd.openxmlformats-officedocument.presentationml.comments+xml"
)
COMMENT_AUTHORS_CONTENT_TYPE = (
    "application/vnd.openxmlformats-officedocument.presentationml.commentAuthors+xml"
)


@dataclass
class CommentAuthor:
    """Represents a comment author."""

    id: int
    name: str
    initials: str = ""
    last_idx: int = 0  # Last comment index by this author
    clr_idx: int = 0  # Color index for author


@dataclass
class Comment:
    """Represents a slide comment."""

    id: int
    author_id: int
    text: str
    date: datetime = field(default_factory=datetime.now)
    x: int = 914400  # Default 1 inch in EMUs
    y: int = 914400


def parse_comment_authors(xml_bytes: bytes) -> list[CommentAuthor]:
    """Parse comment authors from XML.

    Args:
        xml_bytes: The XML content of commentAuthors.xml

    Returns:
        List of CommentAuthor objects
    """
    if not xml_bytes:
        return []

    root = etree.fromstring(xml_bytes)
    authors = []

    for cm_author in root.findall(qn("p:cmAuthor"), namespaces=NAMESPACES):
        author = CommentAuthor(
            id=int(cm_author.get("id", 0)),
            name=cm_author.get("name", ""),
            initials=cm_author.get("initials", ""),
            last_idx=int(cm_author.get("lastIdx", 0)),
            clr_idx=int(cm_author.get("clrIdx", 0)),
        )
        authors.append(author)

    return authors


def create_comment_authors_xml(authors: list[CommentAuthor]) -> bytes:
    """Create comment authors XML.

    Args:
        authors: List of comment authors

    Returns:
        XML content as bytes
    """
    root = etree.Element(
        qn("p:cmAuthorLst"),
        nsmap={
            None: NAMESPACES["p"],
            "r": NAMESPACES["r"],
        },
    )

    for author in authors:
        cm_author = etree.SubElement(root, qn("p:cmAuthor"))
        cm_author.set("id", str(author.id))
        cm_author.set("name", author.name)
        cm_author.set("initials", author.initials)
        cm_author.set("lastIdx", str(author.last_idx))
        cm_author.set("clrIdx", str(author.clr_idx))

    return etree.tostring(
        root,
        xml_declaration=True,
        encoding="UTF-8",
        standalone=True,
    )


def parse_comments(xml_bytes: bytes) -> list[Comment]:
    """Parse comments from XML.

    Args:
        xml_bytes: The XML content of a comments file

    Returns:
        List of Comment objects
    """
    if not xml_bytes:
        return []

    root = etree.fromstring(xml_bytes)
    comments = []

    for cm in root.findall(qn("p:cm"), namespaces=NAMESPACES):
        # Get position
        pos = cm.find(qn("p:pos"), namespaces=NAMESPACES)
        x = int(pos.get("x", 914400)) if pos is not None else 914400
        y = int(pos.get("y", 914400)) if pos is not None else 914400

        # Get text
        text_elem = cm.find(qn("p:text"), namespaces=NAMESPACES)
        text = text_elem.text if text_elem is not None and text_elem.text else ""

        # Parse date
        dt_str = cm.get("dt", "")
        try:
            date = datetime.fromisoformat(dt_str.replace("Z", "+00:00"))
        except (ValueError, TypeError):
            date = datetime.now()

        comment = Comment(
            id=int(cm.get("idx", 0)),
            author_id=int(cm.get("authorId", 0)),
            text=text,
            date=date,
            x=x,
            y=y,
        )
        comments.append(comment)

    return comments


def create_comments_xml(comments: list[Comment]) -> bytes:
    """Create comments XML for a slide.

    Args:
        comments: List of comments

    Returns:
        XML content as bytes
    """
    root = etree.Element(
        qn("p:cmLst"),
        nsmap={
            None: NAMESPACES["p"],
            "r": NAMESPACES["r"],
        },
    )

    for comment in comments:
        cm = etree.SubElement(root, qn("p:cm"))
        cm.set("idx", str(comment.id))
        cm.set("authorId", str(comment.author_id))
        cm.set("dt", comment.date.isoformat())

        # Position
        pos = etree.SubElement(cm, qn("p:pos"))
        pos.set("x", str(comment.x))
        pos.set("y", str(comment.y))

        # Text
        text_elem = etree.SubElement(cm, qn("p:text"))
        text_elem.text = comment.text

    return etree.tostring(
        root,
        xml_declaration=True,
        encoding="UTF-8",
        standalone=True,
    )


def get_author_initials(name: str) -> str:
    """Generate initials from a name.

    Args:
        name: Full name

    Returns:
        Initials (up to 2 characters)
    """
    parts = name.split()
    if len(parts) >= 2:
        return (parts[0][0] + parts[-1][0]).upper()
    elif parts:
        return parts[0][:2].upper()
    return "U"
