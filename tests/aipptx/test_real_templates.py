"""Tests with real templates.

This file contains integration tests that use real PowerPoint templates
to verify the aipptx wrapper works correctly end-to-end.
"""

import json
from pathlib import Path

import pytest

from aipptx import Template
import py2ppt


class TestAWSTemplate:
    """Tests using the AWS corporate template."""

    @pytest.fixture
    def aws_template(self) -> Template | None:
        """Get the AWS template if it exists."""
        path = Path("/Users/user/Documents/py2ppt/AWStempate.pptx")
        if path.exists():
            return Template(path)
        return None

    def test_describe_output(self, aws_template: Template | None) -> None:
        """Test and display template.describe() output."""
        if aws_template is None:
            pytest.skip("AWS template not available")

        layouts = aws_template.describe()

        print("\n" + "=" * 60)
        print("AWS Template Analysis")
        print("=" * 60)

        for layout in layouts:
            print(f"\n{layout['index']}: {layout['name']}")
            print(f"   Type: {layout['type']}")
            print(f"   Description: {layout['description']}")
            print(f"   Best for: {', '.join(layout['best_for'])}")
            print(f"   Placeholders:")
            for name, ph in layout['placeholders'].items():
                print(f"      - {name}: {ph['purpose']} ({ph['position']})")

        print("\n" + "=" * 60)
        print("Theme Colors:")
        print("=" * 60)
        for name, value in aws_template.colors.items():
            print(f"   {name}: {value}")

        print("\n" + "=" * 60)
        print("Theme Fonts:")
        print("=" * 60)
        for role, font in aws_template.fonts.items():
            print(f"   {role}: {font}")

        # Verify structure
        assert len(layouts) > 0
        assert all("name" in l and "placeholders" in l for l in layouts)

    def test_describe_as_text(self, aws_template: Template | None) -> None:
        """Test text description output."""
        if aws_template is None:
            pytest.skip("AWS template not available")

        text = aws_template.describe_as_text()
        print("\n" + "=" * 60)
        print("Template Description (for AI prompts)")
        print("=" * 60)
        print(text)

        assert "Template:" in text
        assert "Layouts:" in text

    def test_layout_recommendations(self, aws_template: Template | None) -> None:
        """Test layout recommendations."""
        if aws_template is None:
            pytest.skip("AWS template not available")

        print("\n" + "=" * 60)
        print("Layout Recommendations")
        print("=" * 60)

        for content_type in ["title", "bullets", "comparison", "image"]:
            recs = aws_template.recommend_layout(content_type)
            print(f"\nFor '{content_type}' content:")
            if recs:
                for rec in recs[:3]:  # Top 3
                    print(f"   {rec['confidence']:.0%} - {rec['name']}: {rec['reason']}")
            else:
                print("   No recommendations")

    def test_create_mcp_vs_skills_presentation(
        self, aws_template: Template | None, tmp_path: Path
    ) -> None:
        """Create the MCP vs Skills presentation using the new API."""
        if aws_template is None:
            pytest.skip("AWS template not available")

        pres = aws_template.create_presentation()

        # Slide 1: Title
        pres.add_title_slide(
            "MCP vs Claude Code Skills",
            "Understanding the Right Tool for the Job"
        )

        # Slide 2: Agenda
        pres.add_content_slide(
            "Topics We'll Cover",
            [
                "What is MCP?",
                "What are Claude Code Skills?",
                "Key Differences",
                "When to Use Each",
                "Decision Framework"
            ]
        )

        # Slide 3: What is MCP?
        pres.add_content_slide(
            "What is MCP?",
            [
                "Model Context Protocol - standard interface for LLM tools",
                "External server process that exposes tools",
                "Language-agnostic (can be any runtime)",
                "Shared across multiple Claude clients",
                "Best for persistent services and integrations"
            ]
        )

        # Slide 4: What are Skills?
        pres.add_content_slide(
            "What are Claude Code Skills?",
            [
                "Prompt-based tool definitions in Claude Code",
                "Written in markdown with instructions",
                "Executed by Claude using existing tools",
                "Local to Claude Code only",
                "Quick to create and iterate on"
            ]
        )

        # Slide 5: Comparison
        pres.add_comparison_slide(
            "MCP vs Skills Comparison",
            "MCP",
            ["Complex setup", "Cross-client", "Any language", "Persistent state"],
            "Skills",
            ["Quick setup", "Claude Code only", "Prompt-based", "Stateless"]
        )

        # Slide 6: When to use MCP
        pres.add_content_slide(
            "When to Use MCP",
            [
                "Need to share tools across clients",
                "Complex logic requiring a specific runtime",
                "Need persistent state between calls",
                "Building for team/enterprise use",
                "Integrating with external services"
            ]
        )

        # Slide 7: When to use Skills
        pres.add_content_slide(
            "When to Use Skills",
            [
                "Quick personal automations",
                "Prompt engineering workflows",
                "Prototyping before building MCP",
                "Simple file operations",
                "Project-specific tasks"
            ]
        )

        # Slide 8: Decision Framework
        pres.add_content_slide(
            "Decision Framework",
            [
                "Start with Skills for rapid prototyping",
                ("Test and refine the approach", 1),
                "Graduate to MCP when needed",
                ("Cross-client requirements", 1),
                ("Complex state management", 1),
                ("Team collaboration", 1),
            ]
        )

        # Slide 9: Section Break
        pres.add_section_slide("Summary")

        # Slide 10: Key Takeaways
        pres.add_content_slide(
            "Key Takeaways",
            [
                "Skills are great for quick, local tasks",
                "MCP is better for shared, complex integrations",
                "They complement each other",
                "Start simple, evolve as needed"
            ]
        )

        # Save
        output_path = tmp_path / "MCP_vs_Skills_test.pptx"
        pres.save(output_path)

        # Verify
        assert output_path.exists()
        assert pres.slide_count == 10

        # Open and verify structure
        loaded = py2ppt.open_presentation(str(output_path))
        assert py2ppt.get_slide_count(loaded) == 10

        print(f"\nCreated presentation with {pres.slide_count} slides")
        print(f"Saved to: {output_path}")

    def test_json_export_for_ai(self, aws_template: Template | None) -> None:
        """Test that describe() output is JSON serializable."""
        if aws_template is None:
            pytest.skip("AWS template not available")

        layouts = aws_template.describe()

        # Should be JSON serializable
        json_str = json.dumps(layouts, indent=2)
        assert len(json_str) > 0

        # Parse back
        parsed = json.loads(json_str)
        assert len(parsed) == len(layouts)

        print("\n" + "=" * 60)
        print("JSON Output (for AI consumption)")
        print("=" * 60)
        print(json_str[:2000] + "..." if len(json_str) > 2000 else json_str)
