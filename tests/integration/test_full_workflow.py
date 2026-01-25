"""End-to-end workflow tests."""

import os

import py2ppt as ppt


class TestAIAgentWorkflow:
    """Test the typical AI agent workflow."""

    def test_complete_presentation_workflow(self, temp_pptx_path):
        """Test creating a complete presentation from scratch."""

        # Step 1: Create presentation
        pres = ppt.create_presentation()

        # Step 2: Inspect available layouts
        layouts = ppt.list_layouts(pres)
        assert len(layouts) > 0
        assert any(layout["name"] for layout in layouts)

        # Step 3: Get theme colors
        colors = ppt.get_theme_colors(pres)
        assert "accent1" in colors

        # Step 4: Create title slide
        ppt.add_slide(pres, layout="Title Slide")
        ppt.set_title(pres, 1, "Q4 Business Review")
        ppt.set_subtitle(pres, 1, "Prepared by Analytics Team")

        # Step 5: Create content slide
        ppt.add_slide(pres, layout="Title and Content")
        ppt.set_title(pres, 2, "Key Metrics")
        ppt.set_body(pres, 2, [
            "Revenue: $4.2M (+20% YoY)",
            "New customers: 1,247",
            "NPS Score: 72"
        ])

        # Step 6: Add another slide with nested bullets
        ppt.add_slide(pres, layout="Title and Content")
        ppt.set_title(pres, 3, "Regional Performance")
        ppt.set_body(pres, 3, [
            "North America",
            "Strong growth in enterprise",
            "EMEA",
            "New market expansion",
        ], levels=[0, 1, 0, 1])

        # Step 7: Verify structure
        assert ppt.get_slide_count(pres) == 3

        # Step 8: Describe slides
        slide1 = ppt.describe_slide(pres, 1)
        assert "placeholders" in slide1

        slide2 = ppt.describe_slide(pres, 2)
        assert "placeholders" in slide2

        # Step 9: Save
        ppt.save_presentation(pres, temp_pptx_path)
        assert os.path.exists(temp_pptx_path)
        assert os.path.getsize(temp_pptx_path) > 0

        # Step 10: Reopen and verify
        pres2 = ppt.open_presentation(temp_pptx_path)
        assert ppt.get_slide_count(pres2) == 3

        slide = pres2.get_slide(1)
        assert slide.get_title() == "Q4 Business Review"


class TestRoundTrip:
    """Test save and reload preserves content."""

    def test_text_roundtrip(self, temp_pptx_path):
        """Test that text content survives save/reload."""
        # Create
        pres = ppt.create_presentation()
        ppt.add_slide(pres, layout="Title Slide")
        ppt.set_title(pres, 1, "Test Title Here")
        ppt.set_subtitle(pres, 1, "Test Subtitle Here")
        ppt.save_presentation(pres, temp_pptx_path)

        # Reload
        pres2 = ppt.open_presentation(temp_pptx_path)
        slide = pres2.get_slide(1)
        assert slide.get_title() == "Test Title Here"
        assert slide.get_subtitle() == "Test Subtitle Here"

    def test_body_roundtrip(self, temp_pptx_path):
        """Test that body content survives save/reload."""
        # Create
        pres = ppt.create_presentation()
        ppt.add_slide(pres, layout="Title and Content")
        ppt.set_title(pres, 1, "Bullets Test")
        bullets = ["First bullet", "Second bullet", "Third bullet"]
        ppt.set_body(pres, 1, bullets)
        ppt.save_presentation(pres, temp_pptx_path)

        # Reload
        pres2 = ppt.open_presentation(temp_pptx_path)
        slide = pres2.get_slide(1)
        body = slide.get_body()
        assert len(body) == 3
        assert body[0] == "First bullet"


class TestTableWorkflow:
    """Test table creation workflow."""

    def test_add_table(self, temp_pptx_path):
        """Test adding a table to a slide."""
        pres = ppt.create_presentation()
        ppt.add_slide(pres, layout="Title and Content")
        ppt.set_title(pres, 1, "Sales Data")

        ppt.add_table(pres, 1, data=[
            ["Region", "Q3", "Q4"],
            ["North", 100, 120],
            ["South", 80, 95],
            ["East", 90, 110],
        ])

        # Save and verify
        ppt.save_presentation(pres, temp_pptx_path)
        assert os.path.exists(temp_pptx_path)

        # Reload and check
        pres2 = ppt.open_presentation(temp_pptx_path)
        info = ppt.describe_slide(pres2, 1)
        # Should have a table shape
        tables = [s for s in info.get("shapes", []) if s.get("type") == "table"]
        assert len(tables) == 1
