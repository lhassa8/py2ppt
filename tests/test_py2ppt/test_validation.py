"""Tests for presentation validation module."""

import pytest
from unittest.mock import MagicMock

from py2ppt.validation import (
    IssueSeverity,
    IssueCategory,
    ValidationIssue,
    ValidationResult,
    validate_slide,
    validate_presentation,
    MAX_BULLETS_PER_SLIDE,
    MAX_WORDS_PER_BULLET,
    MAX_CHARS_PER_BULLET,
)


class TestValidationIssue:
    """Tests for ValidationIssue dataclass."""

    def test_create_issue(self):
        """Test creating a validation issue."""
        issue = ValidationIssue(
            severity=IssueSeverity.WARNING,
            category=IssueCategory.CONTENT,
            slide_number=1,
            message="Too many bullets",
            suggestion="Split across slides",
            rule="too_many_bullets",
        )
        assert issue.severity == IssueSeverity.WARNING
        assert issue.category == IssueCategory.CONTENT
        assert issue.slide_number == 1
        assert issue.message == "Too many bullets"

    def test_to_dict(self):
        """Test to_dict method."""
        issue = ValidationIssue(
            severity=IssueSeverity.ERROR,
            category=IssueCategory.STRUCTURE,
            slide_number=None,
            message="No slides",
            suggestion="Add slides",
            rule="no_slides",
        )
        d = issue.to_dict()
        assert d["severity"] == "error"
        assert d["category"] == "structure"
        assert d["slide_number"] is None
        assert d["message"] == "No slides"

    def test_issue_with_details(self):
        """Test issue with extra details."""
        issue = ValidationIssue(
            severity=IssueSeverity.INFO,
            category=IssueCategory.DESIGN,
            slide_number=3,
            message="Repetitive layout",
            suggestion="Vary layouts",
            rule="repetitive_layout",
            details={"layout": "content", "count": 5},
        )
        d = issue.to_dict()
        assert d["details"]["layout"] == "content"
        assert d["details"]["count"] == 5


class TestValidationResult:
    """Tests for ValidationResult dataclass."""

    def test_empty_result(self):
        """Test empty validation result."""
        result = ValidationResult(is_valid=True, issues=[], score=100.0)
        assert result.is_valid is True
        assert len(result.issues) == 0
        assert result.score == 100.0

    def test_errors_property(self):
        """Test errors property filters correctly."""
        issues = [
            ValidationIssue(IssueSeverity.ERROR, IssueCategory.STRUCTURE, 1, "E1", "S1"),
            ValidationIssue(IssueSeverity.WARNING, IssueCategory.CONTENT, 2, "W1", "S2"),
            ValidationIssue(IssueSeverity.ERROR, IssueCategory.CONTENT, 3, "E2", "S3"),
        ]
        result = ValidationResult(is_valid=False, issues=issues, score=50.0)
        assert len(result.errors) == 2
        assert all(e.severity == IssueSeverity.ERROR for e in result.errors)

    def test_warnings_property(self):
        """Test warnings property filters correctly."""
        issues = [
            ValidationIssue(IssueSeverity.ERROR, IssueCategory.STRUCTURE, 1, "E1", "S1"),
            ValidationIssue(IssueSeverity.WARNING, IssueCategory.CONTENT, 2, "W1", "S2"),
            ValidationIssue(IssueSeverity.WARNING, IssueCategory.DESIGN, 3, "W2", "S3"),
        ]
        result = ValidationResult(is_valid=False, issues=issues, score=66.0)
        assert len(result.warnings) == 2
        assert all(w.severity == IssueSeverity.WARNING for w in result.warnings)

    def test_info_property(self):
        """Test info property filters correctly."""
        issues = [
            ValidationIssue(IssueSeverity.INFO, IssueCategory.CONTENT, 1, "I1", "S1"),
            ValidationIssue(IssueSeverity.WARNING, IssueCategory.CONTENT, 2, "W1", "S2"),
            ValidationIssue(IssueSeverity.INFO, IssueCategory.STRUCTURE, 3, "I2", "S3"),
        ]
        result = ValidationResult(is_valid=True, issues=issues, score=88.0)
        assert len(result.info) == 2
        assert all(i.severity == IssueSeverity.INFO for i in result.info)

    def test_by_slide(self):
        """Test by_slide method."""
        issues = [
            ValidationIssue(IssueSeverity.WARNING, IssueCategory.CONTENT, 1, "W1", "S1"),
            ValidationIssue(IssueSeverity.WARNING, IssueCategory.CONTENT, 2, "W2", "S2"),
            ValidationIssue(IssueSeverity.WARNING, IssueCategory.CONTENT, 1, "W3", "S3"),
        ]
        result = ValidationResult(is_valid=True, issues=issues, score=76.0)
        slide1_issues = result.by_slide(1)
        assert len(slide1_issues) == 2

    def test_by_category(self):
        """Test by_category method."""
        issues = [
            ValidationIssue(IssueSeverity.WARNING, IssueCategory.CONTENT, 1, "W1", "S1"),
            ValidationIssue(IssueSeverity.WARNING, IssueCategory.DESIGN, 2, "W2", "S2"),
            ValidationIssue(IssueSeverity.INFO, IssueCategory.CONTENT, 3, "I1", "S3"),
        ]
        result = ValidationResult(is_valid=True, issues=issues, score=82.0)
        content_issues = result.by_category(IssueCategory.CONTENT)
        assert len(content_issues) == 2

    def test_to_dict(self):
        """Test to_dict method."""
        issues = [
            ValidationIssue(IssueSeverity.WARNING, IssueCategory.CONTENT, 1, "W1", "S1"),
        ]
        result = ValidationResult(is_valid=True, issues=issues, score=92.0)
        d = result.to_dict()
        assert d["is_valid"] is True
        assert d["score"] == 92.0
        assert d["warning_count"] == 1
        assert len(d["issues"]) == 1

    def test_summary(self):
        """Test summary method."""
        issues = [
            ValidationIssue(IssueSeverity.ERROR, IssueCategory.STRUCTURE, 1, "E1", "S1"),
            ValidationIssue(IssueSeverity.WARNING, IssueCategory.CONTENT, 2, "W1", "S2"),
            ValidationIssue(IssueSeverity.INFO, IssueCategory.DESIGN, 3, "I1", "S3"),
        ]
        result = ValidationResult(is_valid=False, issues=issues, score=65.0)
        summary = result.summary()
        assert "NEEDS ATTENTION" in summary
        assert "65" in summary
        assert "1 error" in summary
        assert "1 warning" in summary

    def test_repr(self):
        """Test __repr__ method."""
        result = ValidationResult(is_valid=True, issues=[], score=100.0)
        r = repr(result)
        assert "ValidationResult" in r
        assert "VALID" in r


class TestValidateSlide:
    """Tests for validate_slide function."""

    def test_missing_title_warning(self):
        """Test warning for missing title."""
        slide_info = {
            "slide_number": 1,
            "title": "",
            "content": ["Some content"],
            "layout": "content",
            "notes": "",
        }
        issues = validate_slide(slide_info)
        title_issues = [i for i in issues if i.rule == "missing_title"]
        assert len(title_issues) == 1
        assert title_issues[0].severity == IssueSeverity.WARNING

    def test_missing_title_error_in_strict(self):
        """Test error for missing title in strict mode."""
        slide_info = {
            "slide_number": 1,
            "title": "",
            "content": ["Some content"],
            "layout": "content",
            "notes": "",
        }
        issues = validate_slide(slide_info, strict=True)
        title_issues = [i for i in issues if i.rule == "missing_title"]
        assert len(title_issues) == 1
        assert title_issues[0].severity == IssueSeverity.ERROR

    def test_too_many_bullets(self):
        """Test warning for too many bullets."""
        slide_info = {
            "slide_number": 1,
            "title": "Title",
            "content": [f"Bullet {i}" for i in range(MAX_BULLETS_PER_SLIDE + 3)],
            "layout": "content",
            "notes": "",
        }
        issues = validate_slide(slide_info)
        bullet_issues = [i for i in issues if i.rule == "too_many_bullets"]
        assert len(bullet_issues) == 1

    def test_long_bullet_chars(self):
        """Test warning for bullet exceeding character limit."""
        long_bullet = "x" * (MAX_CHARS_PER_BULLET + 10)
        slide_info = {
            "slide_number": 1,
            "title": "Title",
            "content": [long_bullet],
            "layout": "content",
            "notes": "",
        }
        issues = validate_slide(slide_info)
        long_issues = [i for i in issues if i.rule == "bullet_too_long"]
        assert len(long_issues) == 1

    def test_wordy_bullet(self):
        """Test info for bullet exceeding word limit."""
        wordy_bullet = " ".join(["word"] * (MAX_WORDS_PER_BULLET + 5))
        slide_info = {
            "slide_number": 1,
            "title": "Title",
            "content": [wordy_bullet],
            "layout": "content",
            "notes": "",
        }
        issues = validate_slide(slide_info)
        wordy_issues = [i for i in issues if i.rule == "bullet_wordy"]
        assert len(wordy_issues) == 1
        assert wordy_issues[0].severity == IssueSeverity.INFO

    def test_missing_notes_info(self):
        """Test info for missing speaker notes."""
        slide_info = {
            "slide_number": 2,
            "title": "Content Slide",
            "content": ["Some content"],
            "layout": "content",
            "notes": "",
        }
        issues = validate_slide(slide_info)
        notes_issues = [i for i in issues if i.rule == "missing_notes"]
        assert len(notes_issues) == 1
        assert notes_issues[0].severity == IssueSeverity.INFO

    def test_no_notes_warning_for_title_slide(self):
        """Test no notes warning for title slides."""
        slide_info = {
            "slide_number": 1,
            "title": "Presentation Title",
            "content": [],
            "layout": "title slide",
            "notes": "",
        }
        issues = validate_slide(slide_info)
        notes_issues = [i for i in issues if i.rule == "missing_notes"]
        assert len(notes_issues) == 0

    def test_no_notes_warning_for_blank_slide(self):
        """Test no notes warning for blank slides."""
        slide_info = {
            "slide_number": 3,
            "title": "",
            "content": [],
            "layout": "blank",
            "notes": "",
        }
        issues = validate_slide(slide_info)
        # Blank slides shouldn't trigger notes or title warnings
        notes_issues = [i for i in issues if i.rule == "missing_notes"]
        title_issues = [i for i in issues if i.rule == "missing_title"]
        assert len(notes_issues) == 0
        assert len(title_issues) == 0

    def test_too_much_text(self):
        """Test warning for too much text on slide."""
        slide_info = {
            "slide_number": 1,
            "title": "Title",
            "content": [" ".join(["word"] * 30) for _ in range(4)],  # ~120 words
            "layout": "content",
            "notes": "",
        }
        issues = validate_slide(slide_info)
        text_issues = [i for i in issues if i.rule == "too_much_text"]
        assert len(text_issues) == 1

    def test_empty_content_slide(self):
        """Test info for empty content slide."""
        slide_info = {
            "slide_number": 2,
            "title": "Title Only",
            "content": [],
            "layout": "content",
            "notes": "",
            "has_table": False,
            "has_chart": False,
        }
        issues = validate_slide(slide_info)
        empty_issues = [i for i in issues if i.rule == "empty_content"]
        assert len(empty_issues) == 1

    def test_no_empty_content_for_table_slides(self):
        """Test no empty content warning for slides with tables."""
        slide_info = {
            "slide_number": 2,
            "title": "Data Table",
            "content": [],
            "layout": "content",
            "notes": "",
            "has_table": True,
            "has_chart": False,
        }
        issues = validate_slide(slide_info)
        empty_issues = [i for i in issues if i.rule == "empty_content"]
        assert len(empty_issues) == 0


class TestValidatePresentation:
    """Tests for validate_presentation function."""

    @pytest.fixture
    def mock_presentation(self):
        """Create a mock presentation."""
        pres = MagicMock()
        pres.slide_count = 5
        pres.describe_all_slides.return_value = [
            {
                "slide_number": 1,
                "layout": "title slide",
                "title": "Presentation Title",
                "content": [],
                "notes": "",
                "has_table": False,
                "has_chart": False,
            },
            {
                "slide_number": 2,
                "layout": "content",
                "title": "Introduction",
                "content": ["Point 1", "Point 2"],
                "notes": "Speaker notes",
                "has_table": False,
                "has_chart": False,
            },
            {
                "slide_number": 3,
                "layout": "content",
                "title": "Details",
                "content": ["Detail A", "Detail B"],
                "notes": "More notes",
                "has_table": False,
                "has_chart": False,
            },
            {
                "slide_number": 4,
                "layout": "content",
                "title": "Analysis",
                "content": ["Analysis point"],
                "notes": "",
                "has_table": False,
                "has_chart": False,
            },
            {
                "slide_number": 5,
                "layout": "content",
                "title": "Thank You",
                "content": [],
                "notes": "",
                "has_table": False,
                "has_chart": False,
            },
        ]
        return pres

    def test_valid_presentation(self, mock_presentation):
        """Test validation of a valid presentation."""
        result = validate_presentation(mock_presentation)
        assert result.is_valid is True
        assert result.score > 0

    def test_no_slides_error(self):
        """Test error for presentation with no slides."""
        pres = MagicMock()
        pres.slide_count = 0
        pres.describe_all_slides.return_value = []

        result = validate_presentation(pres)
        assert result.is_valid is False
        assert result.score == 0
        assert any(i.rule == "no_slides" for i in result.errors)

    def test_few_slides_warning(self):
        """Test warning for presentation with only one slide."""
        pres = MagicMock()
        pres.slide_count = 1
        pres.describe_all_slides.return_value = [
            {
                "slide_number": 1,
                "layout": "title slide",
                "title": "Only Slide",
                "content": [],
                "notes": "",
            }
        ]

        result = validate_presentation(pres)
        warnings = [i for i in result.warnings if i.rule == "few_slides"]
        assert len(warnings) == 1

    def test_missing_title_slide_warning(self, mock_presentation):
        """Test warning for missing title slide."""
        # Change first slide to not be a title slide
        mock_presentation.describe_all_slides.return_value[0]["layout"] = "content"

        result = validate_presentation(mock_presentation)
        title_slide_issues = [i for i in result.issues if i.rule == "missing_title_slide"]
        assert len(title_slide_issues) == 1

    def test_missing_closing_info(self):
        """Test info for missing closing slide."""
        pres = MagicMock()
        pres.slide_count = 5
        pres.describe_all_slides.return_value = [
            {"slide_number": i, "layout": "content" if i > 1 else "title slide",
             "title": f"Slide {i}", "content": [], "notes": ""}
            for i in range(1, 6)
        ]
        pres.describe_all_slides.return_value[0]["layout"] = "title slide"

        result = validate_presentation(pres)
        closing_issues = [i for i in result.issues if i.rule == "missing_closing"]
        assert len(closing_issues) == 1

    def test_needs_section_break(self):
        """Test info for too many slides without section break."""
        pres = MagicMock()
        pres.slide_count = 8
        pres.describe_all_slides.return_value = [
            {
                "slide_number": i,
                "layout": "title slide" if i == 1 else "content",
                "title": f"Slide {i}",
                "content": ["Content"],
                "notes": "",
            }
            for i in range(1, 9)
        ]

        result = validate_presentation(pres)
        section_issues = [i for i in result.issues if i.rule == "needs_section_break"]
        # Should trigger after more than 5 content slides
        assert len(section_issues) >= 1

    def test_repetitive_layout(self):
        """Test info for repetitive layouts."""
        pres = MagicMock()
        pres.slide_count = 6
        pres.describe_all_slides.return_value = [
            {
                "slide_number": i,
                "layout": "content",  # All same layout
                "title": f"Slide {i}",
                "content": ["Content"],
                "notes": "",
            }
            for i in range(1, 7)
        ]

        result = validate_presentation(pres)
        layout_issues = [i for i in result.issues if i.rule == "repetitive_layout"]
        assert len(layout_issues) >= 1

    def test_low_layout_variety(self):
        """Test info for low layout variety."""
        pres = MagicMock()
        pres.slide_count = 6
        pres.describe_all_slides.return_value = [
            {
                "slide_number": i,
                "layout": "content",
                "title": f"Slide {i}",
                "content": [],
                "notes": "",
            }
            for i in range(1, 7)
        ]

        result = validate_presentation(pres)
        variety_issues = [i for i in result.issues if i.rule == "low_layout_variety"]
        assert len(variety_issues) == 1

    def test_strict_mode(self, mock_presentation):
        """Test strict mode treats warnings as failures."""
        # Add a slide without title
        mock_presentation.describe_all_slides.return_value[3]["title"] = ""

        result_normal = validate_presentation(mock_presentation, strict=False)
        result_strict = validate_presentation(mock_presentation, strict=True)

        # Normal mode: warnings don't affect validity
        assert len(result_normal.warnings) > 0

        # Strict mode: warnings affect validity
        assert result_strict.is_valid is False

    def test_score_calculation(self, mock_presentation):
        """Test score is calculated correctly."""
        result = validate_presentation(mock_presentation)
        # With a mostly valid presentation, score should be high
        assert result.score >= 70

    def test_score_decreases_with_issues(self):
        """Test score decreases with issues."""
        pres = MagicMock()
        pres.slide_count = 3
        # Create slides with problems
        pres.describe_all_slides.return_value = [
            {
                "slide_number": 1,
                "layout": "content",  # Not a title slide
                "title": "",  # Missing title
                "content": [f"Bullet {i}" for i in range(10)],  # Too many bullets
                "notes": "",
            },
            {
                "slide_number": 2,
                "layout": "content",
                "title": "",
                "content": ["x" * 150],  # Long bullet
                "notes": "",
            },
            {
                "slide_number": 3,
                "layout": "content",
                "title": "",
                "content": [],
                "notes": "",
            },
        ]

        result = validate_presentation(pres)
        # Score should be significantly reduced
        assert result.score < 80


class TestEnums:
    """Tests for enum values."""

    def test_issue_severity_values(self):
        """Test IssueSeverity enum values."""
        assert IssueSeverity.ERROR.value == "error"
        assert IssueSeverity.WARNING.value == "warning"
        assert IssueSeverity.INFO.value == "info"

    def test_issue_category_values(self):
        """Test IssueCategory enum values."""
        assert IssueCategory.CONTENT.value == "content"
        assert IssueCategory.STRUCTURE.value == "structure"
        assert IssueCategory.DESIGN.value == "design"
        assert IssueCategory.ACCESSIBILITY.value == "accessibility"
