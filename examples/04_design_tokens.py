#!/usr/bin/env python3
"""Example: Using design tokens for brand consistency.

Design tokens provide a way to define and enforce brand guidelines
across presentations. This example shows how to create, use, and
validate with design tokens.
"""

import py2ppt as ppt
from py2ppt.template.tokens import create_tokens, save_tokens, load_tokens


def main():
    print("=" * 60)
    print("Design Tokens Example")
    print("=" * 60)

    # Step 1: Define brand tokens
    print("\n[1] Defining brand tokens...")
    tokens = create_tokens({
        "colors": {
            "brand-primary": "#0066CC",
            "brand-secondary": "#FF6600",
            "brand-accent": "#00AA55",
            "text-dark": "#333333",
            "text-light": "#666666",
            "background": "#FFFFFF",
            "success": "#00AA55",
            "warning": "#FFAA00",
            "error": "#CC0000",
        },
        "fonts": {
            "heading": {"family": "Arial Black", "size": "36pt"},
            "subheading": {"family": "Arial", "size": "24pt", "weight": "bold"},
            "body": {"family": "Arial", "size": "18pt"},
            "caption": {"family": "Arial", "size": "12pt"},
        },
        "metadata": {
            "brand": "ACME Corp",
            "version": "1.0",
        }
    })

    print(f"   Colors defined: {list(tokens.colors.keys())}")
    print(f"   Fonts defined: {list(tokens.fonts.keys())}")

    # Step 2: Save tokens for reuse
    tokens_path = "brand_tokens.json"
    print(f"\n[2] Saving tokens to {tokens_path}...")
    save_tokens(tokens, tokens_path)

    # Step 3: Load tokens (simulating a new session)
    print(f"\n[3] Loading tokens from {tokens_path}...")
    loaded_tokens = load_tokens(tokens_path)
    print(f"   Primary color: {loaded_tokens.get_color('brand-primary')}")
    print(f"   Heading font: {loaded_tokens.get_font('heading').family}")

    # Step 4: Use tokens in a presentation
    print("\n[4] Creating presentation with brand tokens...")
    pres = ppt.create_presentation()

    # Add title slide with brand colors
    ppt.add_slide(pres, layout="Title Slide")
    ppt.set_title(pres, 1, "ACME Corp Quarterly Review")
    ppt.set_subtitle(pres, 1, "Building the future, today")

    # Apply brand styling
    ppt.set_text_style(pres, 1, "title",
        font=loaded_tokens.get_font("heading").family,
        color=loaded_tokens.get_color("brand-primary"),
        bold=True
    )

    # Content slide
    ppt.add_slide(pres, layout="Title and Content")
    ppt.set_title(pres, 2, "Key Achievements")
    ppt.set_body(pres, 2, [
        "Revenue exceeded targets by 15%",
        "Launched 3 new products",
        "Expanded to 5 new markets",
        "Customer satisfaction at all-time high",
    ], color=loaded_tokens.get_color("text-dark"))

    # Save presentation
    output_path = "branded_presentation.pptx"
    ppt.save_presentation(pres, output_path)
    print(f"\n[5] Saved branded presentation to {output_path}")

    # Cleanup
    import os
    os.remove(tokens_path)
    print("\nDesign tokens example complete!")


if __name__ == "__main__":
    main()
