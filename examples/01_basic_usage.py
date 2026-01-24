#!/usr/bin/env python3
"""Basic usage example for py2ppt.

This example demonstrates the simplest way to create a PowerPoint
presentation using py2ppt.
"""

import py2ppt as ppt


def main():
    # Create a new blank presentation
    print("Creating presentation...")
    pres = ppt.create_presentation()

    # Check available layouts
    layouts = ppt.list_layouts(pres)
    print(f"Available layouts: {[l['name'] for l in layouts]}")

    # Add a title slide
    print("Adding title slide...")
    ppt.add_slide(pres, layout="Title Slide")
    ppt.set_title(pres, 1, "Hello, World!")
    ppt.set_subtitle(pres, 1, "My first py2ppt presentation")

    # Add a content slide
    print("Adding content slide...")
    ppt.add_slide(pres, layout="Title and Content")
    ppt.set_title(pres, 2, "Key Features")
    ppt.set_body(pres, 2, [
        "Simple, clean API",
        "Designed for AI agents",
        "No magic numbers",
        "Works with any template",
    ])

    # Add another slide with nested bullets
    print("Adding slide with nested bullets...")
    ppt.add_slide(pres, layout="Title and Content")
    ppt.set_title(pres, 3, "Getting Started")
    ppt.set_body(pres, 3, [
        "Install py2ppt",
        "pip install py2ppt",
        "Import and use",
        "import py2ppt as ppt",
    ], levels=[0, 1, 0, 1])

    # Save the presentation
    output_path = "basic_example.pptx"
    print(f"Saving to {output_path}...")
    ppt.save_presentation(pres, output_path)

    print(f"Done! Created presentation with {ppt.get_slide_count(pres)} slides.")
    print(f"Open {output_path} to view the result.")


if __name__ == "__main__":
    main()
