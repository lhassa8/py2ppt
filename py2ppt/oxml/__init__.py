"""Open XML handling for PowerPoint files.

This module provides low-level XML manipulation for .pptx files.
"""

from .ns import nsmap, qn
from .package import Package
from .presentation import PresentationPart
from .slide import SlidePart
from .layout import SlideLayoutPart
from .master import SlideMasterPart
from .theme import ThemePart
from .shapes import ShapeTree, Shape, TextFrame
from .text import Paragraph, Run

__all__ = [
    "nsmap",
    "qn",
    "Package",
    "PresentationPart",
    "SlidePart",
    "SlideLayoutPart",
    "SlideMasterPart",
    "ThemePart",
    "ShapeTree",
    "Shape",
    "TextFrame",
    "Paragraph",
    "Run",
]
