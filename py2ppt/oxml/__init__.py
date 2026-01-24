"""Open XML handling for PowerPoint files.

This module provides low-level XML manipulation for .pptx files.
"""

from .layout import SlideLayoutPart
from .master import SlideMasterPart
from .ns import nsmap, qn
from .package import Package
from .presentation import PresentationPart
from .shapes import Shape, ShapeTree, TextFrame
from .slide import SlidePart
from .text import Paragraph, Run
from .theme import ThemePart

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
