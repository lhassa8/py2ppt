"""Tests for find and replace functionality."""

import py2ppt as ppt


class TestFindText:
    """Tests for find_text function."""

    def test_find_text_in_title(self):
        """Test finding text in slide titles."""
        pres = ppt.create_presentation()
        ppt.add_slide(pres, layout="Title and Content")
        ppt.set_title(pres, 1, "Welcome to Testing")

        results = ppt.find_text(pres, "Testing")

        assert len(results) >= 1
        assert any(r["slide"] == 1 for r in results)

    def test_find_text_case_insensitive(self):
        """Test case-insensitive search."""
        pres = ppt.create_presentation()
        ppt.add_slide(pres, layout="Title and Content")
        ppt.set_title(pres, 1, "Hello World")

        results = ppt.find_text(pres, "hello", case_sensitive=False)

        assert len(results) >= 1

    def test_find_text_case_sensitive(self):
        """Test case-sensitive search."""
        pres = ppt.create_presentation()
        ppt.add_slide(pres, layout="Title and Content")
        ppt.set_title(pres, 1, "Hello World")

        # Should not find lowercase
        results = ppt.find_text(pres, "hello", case_sensitive=True)
        assert len(results) == 0

        # Should find correct case
        results = ppt.find_text(pres, "Hello", case_sensitive=True)
        assert len(results) >= 1

    def test_find_text_not_found(self):
        """Test searching for non-existent text."""
        pres = ppt.create_presentation()
        ppt.add_slide(pres, layout="Title and Content")
        ppt.set_title(pres, 1, "Hello World")

        results = ppt.find_text(pres, "xyz123nonexistent")

        assert len(results) == 0


class TestReplaceText:
    """Tests for replace_text function."""

    def test_replace_text_single(self):
        """Test replacing text in presentation."""
        pres = ppt.create_presentation()
        ppt.add_slide(pres, layout="Title and Content")
        ppt.set_title(pres, 1, "Hello World")

        count = ppt.replace_text(pres, "World", "Universe")

        assert count >= 1

    def test_replace_text_case_insensitive(self):
        """Test case-insensitive replace."""
        pres = ppt.create_presentation()
        ppt.add_slide(pres, layout="Title and Content")
        ppt.set_title(pres, 1, "Hello World")

        count = ppt.replace_text(pres, "world", "Universe", case_sensitive=False)

        assert count >= 1


class TestReplaceAll:
    """Tests for replace_all function."""

    def test_replace_all_multiple(self):
        """Test replacing multiple strings at once."""
        pres = ppt.create_presentation()
        ppt.add_slide(pres, layout="Title and Content")
        ppt.set_title(pres, 1, "Hello World")
        ppt.add_slide(pres, layout="Title and Content")
        ppt.set_title(pres, 2, "Goodbye World")

        replacements = {
            "Hello": "Hi",
            "Goodbye": "Farewell",
        }
        counts = ppt.replace_all(pres, replacements)

        assert isinstance(counts, dict)
