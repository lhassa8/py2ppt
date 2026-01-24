"""Pytest configuration and fixtures."""

import os
import tempfile
from pathlib import Path

import pytest

import py2ppt as ppt


@pytest.fixture
def blank_presentation():
    """Create a blank presentation."""
    return ppt.create_presentation()


@pytest.fixture
def presentation_with_slides():
    """Create a presentation with a few slides."""
    pres = ppt.create_presentation()
    ppt.add_slide(pres, layout="Title Slide")
    ppt.set_title(pres, 1, "Test Presentation")
    ppt.set_subtitle(pres, 1, "A test subtitle")

    ppt.add_slide(pres, layout="Title and Content")
    ppt.set_title(pres, 2, "Key Points")
    ppt.set_body(pres, 2, ["Point 1", "Point 2", "Point 3"])

    return pres


@pytest.fixture
def temp_pptx_path():
    """Create a temporary file path for saving presentations."""
    with tempfile.NamedTemporaryFile(suffix=".pptx", delete=False) as f:
        path = f.name

    yield path

    # Cleanup
    if os.path.exists(path):
        os.unlink(path)


@pytest.fixture
def fixtures_dir():
    """Get the fixtures directory path."""
    return Path(__file__).parent / "fixtures"
