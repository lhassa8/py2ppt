"""Demo: AI-Friendly PowerPoint Wrapper (aipptx)

This example shows how to use the aipptx wrapper to create
professional PowerPoint presentations from corporate templates
with AI-friendly, semantic APIs.
"""

from pathlib import Path

from aipptx import Template

# Path to template (use your corporate template)
TEMPLATE_PATH = Path(__file__).parent.parent / "AWStempate.pptx"
OUTPUT_PATH = Path(__file__).parent.parent / "demo_output.pptx"


def main():
    # ==========================================
    # 1. Load and Analyze Template
    # ==========================================
    print("Loading template...")
    template = Template(TEMPLATE_PATH)

    # Get AI-friendly descriptions of layouts
    print("\n=== Available Layouts ===")
    for layout in template.describe():
        print(f"\n{layout['index']}: {layout['name']} ({layout['type']})")
        print(f"   Best for: {', '.join(layout['best_for'])}")
        placeholders = list(layout['placeholders'].keys())
        print(f"   Placeholders: {', '.join(placeholders)}")

    # Get theme colors for brand consistency
    print("\n=== Theme Colors ===")
    for name, value in list(template.colors.items())[:6]:
        print(f"   {name}: {value}")

    # Get theme fonts
    print("\n=== Theme Fonts ===")
    for role, font in template.fonts.items():
        print(f"   {role}: {font}")

    # ==========================================
    # 2. Get Layout Recommendations
    # ==========================================
    print("\n=== Layout Recommendations ===")
    for content_type in ["title", "bullets", "comparison"]:
        recs = template.recommend_layout(content_type)
        if recs:
            best = recs[0]
            print(f"For '{content_type}': {best['name']} ({best['confidence']:.0%})")

    # ==========================================
    # 3. Create Presentation
    # ==========================================
    print("\n=== Creating Presentation ===")
    pres = template.create_presentation()

    # Title slide
    pres.add_title_slide(
        "Q4 Business Review",
        "January 2025"
    )
    print("Added: Title slide")

    # Agenda slide with bullets
    pres.add_content_slide(
        "Topics",
        [
            "Financial Performance",
            "Customer Metrics",
            "Product Updates",
            "Next Quarter Goals"
        ]
    )
    print("Added: Agenda slide")

    # Content with nested bullets
    pres.add_content_slide(
        "Financial Highlights",
        [
            "Revenue exceeded targets by 15%",
            "Operating margin improved",
            ("Cost savings from automation", 1),
            ("Reduced overhead", 1),
            "Cash flow remains strong"
        ]
    )
    print("Added: Financial highlights")

    # Comparison slide
    pres.add_comparison_slide(
        "Year Over Year",
        "2023",
        ["Revenue: $100M", "Customers: 5,000", "NPS: 45"],
        "2024",
        ["Revenue: $120M", "Customers: 7,500", "NPS: 62"]
    )
    print("Added: Comparison slide")

    # Section break
    pres.add_section_slide("Next Steps")
    print("Added: Section slide")

    # Action items
    pres.add_content_slide(
        "Q1 Priorities",
        [
            "Complete platform migration",
            "Launch mobile app v2.0",
            "Expand to 3 new markets",
            "Hire 50 engineers"
        ]
    )
    print("Added: Priorities slide")

    # Save
    pres.save(OUTPUT_PATH)
    print(f"\n=== Saved to {OUTPUT_PATH} ===")
    print(f"Total slides: {pres.slide_count}")


if __name__ == "__main__":
    if not TEMPLATE_PATH.exists():
        print(f"Template not found: {TEMPLATE_PATH}")
        print("Please provide a template file path")
    else:
        main()
