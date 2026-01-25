"""Tests for speaker notes functionality."""

import py2ppt as ppt


class TestSetNotes:
    """Tests for set_notes function."""

    def test_set_notes(self):
        """Test setting speaker notes on a slide."""
        pres = ppt.create_presentation()
        ppt.add_slide(pres, layout="Title and Content")

        ppt.set_notes(pres, 1, "This is a speaker note.")
        notes = ppt.get_notes(pres, 1)

        assert "This is a speaker note" in notes

    def test_set_notes_multiline(self):
        """Test setting multiline speaker notes."""
        pres = ppt.create_presentation()
        ppt.add_slide(pres, layout="Title and Content")

        text = "First point\nSecond point\nThird point"
        ppt.set_notes(pres, 1, text)
        notes = ppt.get_notes(pres, 1)

        assert "First point" in notes
        assert "Second point" in notes
        assert "Third point" in notes


class TestGetNotes:
    """Tests for get_notes function."""

    def test_get_notes_empty(self):
        """Test getting notes from slide without notes."""
        pres = ppt.create_presentation()
        ppt.add_slide(pres, layout="Title and Content")

        notes = ppt.get_notes(pres, 1)
        # Empty notes or no notes is valid
        assert notes == "" or notes is not None


class TestAppendNotes:
    """Tests for append_notes function."""

    def test_append_notes(self):
        """Test appending to existing notes."""
        pres = ppt.create_presentation()
        ppt.add_slide(pres, layout="Title and Content")

        ppt.set_notes(pres, 1, "First note")
        ppt.append_notes(pres, 1, "Additional note")
        notes = ppt.get_notes(pres, 1)

        assert "First note" in notes
        assert "Additional note" in notes

    def test_append_notes_empty_start(self):
        """Test appending notes to slide without existing notes."""
        pres = ppt.create_presentation()
        ppt.add_slide(pres, layout="Title and Content")

        ppt.append_notes(pres, 1, "New note")
        notes = ppt.get_notes(pres, 1)

        assert "New note" in notes
