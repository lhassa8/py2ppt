#!/usr/bin/env python3
"""Example: Creating a presentation from a template.

This example shows how to use an existing PowerPoint template
and populate it with content while preserving the template's
styling and branding.
"""

import py2ppt as ppt


def main():
    # For this example, we'll create a "template" first
    # In real usage, you'd use an actual corporate template
    template_path = "sample_template.pptx"

    # Create a simple template
    print("Creating sample template...")
    template = ppt.create_presentation()
    ppt.add_slide(template, layout="Title Slide")
    ppt.save_presentation(template, template_path)

    # Now use the template
    print(f"Opening template: {template_path}")
    pres = ppt.create_presentation(template=template_path)

    # Inspect the template
    print("\nTemplate Analysis:")
    layouts = ppt.list_layouts(pres)
    for layout in layouts:
        print(f"  - {layout['name']}: {layout['placeholders']}")

    # Get theme colors
    colors = ppt.get_theme_colors(pres)
    print("\nTheme Colors:")
    for name, color in list(colors.items())[:6]:
        print(f"  - {name}: {color}")

    # Get theme fonts
    fonts = ppt.get_theme_fonts(pres)
    print("\nTheme Fonts:")
    for role, font in fonts.items():
        print(f"  - {role}: {font}")

    # Create slides using the template's styling
    print("\nCreating slides...")

    ppt.add_slide(pres, layout="Title Slide")
    ppt.set_title(pres, 1, "Q4 Business Review")
    ppt.set_subtitle(pres, 1, "Confidential - Internal Use Only")

    ppt.add_slide(pres, layout="Title and Content")
    ppt.set_title(pres, 2, "Executive Summary")
    ppt.set_body(pres, 2, [
        "Record revenue quarter: $4.2M",
        "Customer base grew 25%",
        "Launched in 3 new markets",
        "NPS improved to 72 (+8 pts)",
    ])

    ppt.add_slide(pres, layout="Title and Content")
    ppt.set_title(pres, 3, "Financial Highlights")
    ppt.add_table(pres, 3, data=[
        ["Metric", "Q3", "Q4", "Change"],
        ["Revenue", "$3.5M", "$4.2M", "+20%"],
        ["Gross Margin", "68%", "71%", "+3pp"],
        ["Operating Expense", "$2.1M", "$2.3M", "+10%"],
    ])

    # Save
    output_path = "from_template.pptx"
    ppt.save_presentation(pres, output_path)
    print(f"\nSaved presentation to {output_path}")

    # Clean up template
    import os
    os.remove(template_path)


if __name__ == "__main__":
    main()
