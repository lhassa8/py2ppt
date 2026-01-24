#!/usr/bin/env python3
"""Example: How an AI agent would use py2ppt.

This example demonstrates the workflow an AI agent (LLM) would
follow when creating a presentation. The key steps are:

1. Inspect the template to understand available layouts and placeholders
2. Use that information to make informed decisions
3. Handle errors gracefully with self-correction
4. Create content that fits the template constraints
"""

import py2ppt as ppt
from py2ppt.utils.errors import LayoutNotFoundError


def ai_agent_workflow():
    """Simulate an AI agent creating a presentation."""

    print("=" * 60)
    print("AI Agent: Starting presentation creation")
    print("=" * 60)

    # Step 1: Create presentation and inspect capabilities
    print("\n[Agent] Creating presentation and inspecting template...")
    pres = ppt.create_presentation()

    # Agent inspects available layouts
    layouts = ppt.list_layouts(pres)
    print(f"[Agent] Found {len(layouts)} layouts:")
    for layout in layouts:
        print(f"        - {layout['name']}: {layout['placeholders']}")

    # Agent inspects theme for consistent styling
    colors = ppt.get_theme_colors(pres)
    fonts = ppt.get_theme_fonts(pres)
    print(f"\n[Agent] Theme colors available: {list(colors.keys())}")
    print(f"[Agent] Theme fonts: heading='{fonts['heading']}', body='{fonts['body']}'")

    # Step 2: Agent plans the presentation structure
    print("\n[Agent] Planning presentation structure...")
    slides_plan = [
        {"layout": "Title Slide", "purpose": "Cover slide"},
        {"layout": "Title and Content", "purpose": "Agenda"},
        {"layout": "Title and Content", "purpose": "Key metrics"},
        {"layout": "Title and Content", "purpose": "Details"},
        {"layout": "Title and Content", "purpose": "Next steps"},
    ]
    print(f"[Agent] Planned {len(slides_plan)} slides")

    # Step 3: Agent creates slides with error handling
    print("\n[Agent] Creating slides...")

    for i, plan in enumerate(slides_plan, 1):
        try:
            ppt.add_slide(pres, layout=plan["layout"])
            print(f"[Agent] Created slide {i}: {plan['purpose']}")
        except LayoutNotFoundError as e:
            # Agent self-corrects by checking available layouts
            print(f"[Agent] Layout error: {e}")
            print(f"[Agent] Falling back to default layout...")
            ppt.add_slide(pres, layout=0)

    # Step 4: Agent populates content
    print("\n[Agent] Populating slide content...")

    # Slide 1: Title
    ppt.set_title(pres, 1, "Q4 Business Review")
    ppt.set_subtitle(pres, 1, "Analytics Team | December 2024")
    print("[Agent] Slide 1: Added title and subtitle")

    # Slide 2: Agenda
    ppt.set_title(pres, 2, "Agenda")
    ppt.set_body(pres, 2, [
        "Financial Performance",
        "Customer Metrics",
        "Product Updates",
        "2025 Outlook",
    ])
    print("[Agent] Slide 2: Added agenda items")

    # Slide 3: Key Metrics with formatting
    ppt.set_title(pres, 3, "Key Metrics")
    ppt.set_body(pres, 3, [
        "Revenue: $4.2M (+20% YoY)",
        "Gross Margin: 71% (+3pp)",
        "Active Users: 125K (+45%)",
        "NPS: 72 (+8 points)",
    ])
    print("[Agent] Slide 3: Added key metrics")

    # Slide 4: Details with nested bullets
    ppt.set_title(pres, 4, "Regional Performance")
    ppt.set_body(pres, 4, [
        "North America",
        "Enterprise segment grew 35%",
        "SMB stable at $1.2M",
        "EMEA",
        "New markets: Germany, France",
        "Pipeline doubled vs Q3",
    ], levels=[0, 1, 1, 0, 1, 1])
    print("[Agent] Slide 4: Added regional details with nested bullets")

    # Slide 5: Next steps
    ppt.set_title(pres, 5, "Next Steps")
    ppt.set_body(pres, 5, [
        "Finalize 2025 budget by Dec 15",
        "Launch APAC expansion in Q1",
        "Complete product roadmap review",
    ])
    print("[Agent] Slide 5: Added next steps")

    # Step 5: Agent verifies the result
    print("\n[Agent] Verifying presentation...")
    total_slides = ppt.get_slide_count(pres)
    print(f"[Agent] Total slides created: {total_slides}")

    for i in range(1, total_slides + 1):
        info = ppt.describe_slide(pres, i)
        title = info.get("placeholders", {}).get("title", "")
        if isinstance(title, str):
            title = title[:40] + "..." if len(title) > 40 else title
        print(f"[Agent] Slide {i}: {title}")

    # Step 6: Save the result
    output_path = "ai_generated.pptx"
    ppt.save_presentation(pres, output_path)
    print(f"\n[Agent] Saved presentation to {output_path}")
    print("=" * 60)
    print("AI Agent: Task completed successfully")
    print("=" * 60)

    return pres


def demonstrate_error_handling():
    """Show how the API helps AI agents self-correct."""

    print("\n" + "=" * 60)
    print("Demonstrating Error Handling for AI Agents")
    print("=" * 60)

    pres = ppt.create_presentation()

    # Try to use an invalid layout
    print("\n[Agent] Trying to add slide with invalid layout 'TwoColumn'...")
    try:
        ppt.add_slide(pres, layout="TwoColumn")
    except LayoutNotFoundError as e:
        print(f"[Agent] Error: {e}")
        # The error includes suggestions
        print("[Agent] Self-correcting by trying 'Two Content' instead...")
        ppt.add_slide(pres, layout="Two Content")
        print("[Agent] Success!")

    print("\n[Agent] Error handling demonstration complete")


if __name__ == "__main__":
    ai_agent_workflow()
    demonstrate_error_handling()
