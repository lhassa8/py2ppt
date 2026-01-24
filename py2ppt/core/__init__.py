"""Core abstractions for py2ppt."""

from .presentation import Presentation
from .slide import Slide
from .placeholder import PlaceholderMap, normalize_placeholder_name

__all__ = [
    "Presentation",
    "Slide",
    "PlaceholderMap",
    "normalize_placeholder_name",
]
