"""Tests for slide comments functionality."""

import py2ppt as ppt


class TestAddComment:
    """Tests for add_comment function."""

    def test_add_comment(self):
        """Test adding a comment to a slide."""
        pres = ppt.create_presentation()
        ppt.add_slide(pres, layout="Blank")

        comment_id = ppt.add_comment(pres, 1, "Review needed", "John Smith")

        assert comment_id > 0

    def test_add_comment_with_position(self):
        """Test adding a comment at a specific position."""
        pres = ppt.create_presentation()
        ppt.add_slide(pres, layout="Blank")

        comment_id = ppt.add_comment(
            pres, 1, "Check this area", "Jane Doe", left="2in", top="3in"
        )

        assert comment_id > 0

    def test_add_multiple_comments(self):
        """Test adding multiple comments to a slide."""
        pres = ppt.create_presentation()
        ppt.add_slide(pres, layout="Blank")

        id1 = ppt.add_comment(pres, 1, "First comment", "Author 1")
        id2 = ppt.add_comment(pres, 1, "Second comment", "Author 2")

        # IDs should be unique
        assert id1 > 0
        assert id2 > 0

    def test_add_comment_same_author(self):
        """Test adding multiple comments from same author."""
        pres = ppt.create_presentation()
        ppt.add_slide(pres, layout="Blank")

        id1 = ppt.add_comment(pres, 1, "Comment 1", "John Smith")
        id2 = ppt.add_comment(pres, 1, "Comment 2", "John Smith")

        assert id1 > 0
        assert id2 > 0
        assert id2 > id1  # Second comment should have higher ID

    def test_add_comment_default_author(self):
        """Test adding comment with default author."""
        pres = ppt.create_presentation()
        ppt.add_slide(pres, layout="Blank")

        comment_id = ppt.add_comment(pres, 1, "Anonymous comment")

        assert comment_id > 0


class TestGetComments:
    """Tests for get_comments function."""

    def test_get_comments(self):
        """Test getting comments from a slide."""
        pres = ppt.create_presentation()
        ppt.add_slide(pres, layout="Blank")
        ppt.add_comment(pres, 1, "Test comment", "Tester")

        comments = ppt.get_comments(pres, 1)

        assert len(comments) >= 1
        assert comments[0]["text"] == "Test comment"
        assert comments[0]["author"] == "Tester"

    def test_get_comments_empty(self):
        """Test getting comments from slide with no comments."""
        pres = ppt.create_presentation()
        ppt.add_slide(pres, layout="Blank")

        comments = ppt.get_comments(pres, 1)

        assert isinstance(comments, list)
        assert len(comments) == 0

    def test_get_comments_returns_all_fields(self):
        """Test that get_comments returns all expected fields."""
        pres = ppt.create_presentation()
        ppt.add_slide(pres, layout="Blank")
        ppt.add_comment(pres, 1, "Field test", "Author")

        comments = ppt.get_comments(pres, 1)

        assert len(comments) >= 1
        comment = comments[0]
        assert "id" in comment
        assert "author" in comment
        assert "text" in comment
        assert "date" in comment
        assert "left" in comment
        assert "top" in comment

    def test_get_comments_invalid_slide(self):
        """Test getting comments from invalid slide."""
        pres = ppt.create_presentation()
        ppt.add_slide(pres, layout="Blank")

        comments = ppt.get_comments(pres, 999)

        assert comments == []


class TestDeleteComment:
    """Tests for delete_comment function."""

    def test_delete_comment(self):
        """Test deleting a comment."""
        pres = ppt.create_presentation()
        ppt.add_slide(pres, layout="Blank")
        comment_id = ppt.add_comment(pres, 1, "To be deleted", "Author")

        result = ppt.delete_comment(pres, 1, comment_id)

        assert result is True
        comments = ppt.get_comments(pres, 1)
        assert len(comments) == 0

    def test_delete_nonexistent_comment(self):
        """Test deleting a comment that doesn't exist."""
        pres = ppt.create_presentation()
        ppt.add_slide(pres, layout="Blank")

        result = ppt.delete_comment(pres, 1, 9999)

        assert result is False

    def test_delete_one_of_multiple_comments(self):
        """Test deleting one comment preserves others."""
        pres = ppt.create_presentation()
        ppt.add_slide(pres, layout="Blank")
        ppt.add_comment(pres, 1, "Keep this", "Author")
        id_to_delete = ppt.add_comment(pres, 1, "Delete this", "Author")

        ppt.delete_comment(pres, 1, id_to_delete)
        comments = ppt.get_comments(pres, 1)

        assert len(comments) == 1
        assert comments[0]["text"] == "Keep this"

    def test_delete_comment_invalid_slide(self):
        """Test deleting comment from invalid slide."""
        pres = ppt.create_presentation()
        ppt.add_slide(pres, layout="Blank")

        result = ppt.delete_comment(pres, 999, 1)

        assert result is False
