#!/usr/bin/env python3
"""Example: Validating presentations against style guides.

This example shows how to define style guide rules and validate
presentations to ensure they meet brand and design standards.
"""

import py2ppt as ppt
from py2ppt.validation import create_style_guide, validate


def main():
    print("=" * 60)
    print("Style Guide Validation Example")
    print("=" * 60)

    # Step 1: Define style guide rules
    print("\n[1] Defining style guide rules...")
    rules = create_style_guide({
        "max_bullet_points": 5,
        "max_words_per_bullet": 10,
        "max_slides": 20,
        "max_title_length": 50,
        "forbidden_fonts": ["Comic Sans MS", "Papyrus", "Wingdings"],
        "min_font_size": 14,
    })

    print("   Rules defined:")
    print(f"   - Max bullet points per slide: {rules.max_bullet_points}")
    print(f"   - Max words per bullet: {rules.max_words_per_bullet}")
    print(f"   - Max slides: {rules.max_slides}")
    print(f"   - Min font size: {rules.min_font_size}pt")
    print(f"   - Forbidden fonts: {rules.forbidden_fonts}")

    # Step 2: Create a presentation that violates some rules
    print("\n[2] Creating test presentation with intentional violations...")
    pres = ppt.create_presentation()

    # Slide 1: Good slide
    ppt.add_slide(pres, layout="Title Slide")
    ppt.set_title(pres, 1, "Valid Presentation")
    ppt.set_subtitle(pres, 1, "This slide is fine")

    # Slide 2: Too many bullets (violation)
    ppt.add_slide(pres, layout="Title and Content")
    ppt.set_title(pres, 2, "Too Many Bullets")
    ppt.set_body(pres, 2, [
        "Point 1",
        "Point 2",
        "Point 3",
        "Point 4",
        "Point 5",
        "Point 6 - this exceeds the limit",
        "Point 7 - way too many",
    ])

    # Slide 3: Bullet with too many words (violation)
    ppt.add_slide(pres, layout="Title and Content")
    ppt.set_title(pres, 3, "Wordy Bullets")
    ppt.set_body(pres, 3, [
        "This is a good bullet point",
        "This bullet point has way too many words and should be flagged by the validator for being too long",
        "Short and sweet",
    ])

    # Slide 4: Title too long (violation)
    ppt.add_slide(pres, layout="Title and Content")
    long_title = "This is an extremely long title that definitely exceeds the maximum character limit"
    ppt.set_title(pres, 4, long_title)
    ppt.set_body(pres, 4, ["Content here"])

    print(f"   Created {ppt.get_slide_count(pres)} slides")

    # Step 3: Validate the presentation
    print("\n[3] Validating presentation against style guide...")
    issues = validate(pres, rules)

    if not issues:
        print("   No issues found!")
    else:
        print(f"   Found {len(issues)} issue(s):")
        for issue in issues:
            slide_str = f"Slide {issue.slide}" if issue.slide else "Presentation"
            severity = issue.severity.upper()
            print(f"\n   [{severity}] {slide_str}")
            print(f"   Rule: {issue.rule}")
            print(f"   Message: {issue.message}")

    # Step 4: Create a compliant presentation
    print("\n[4] Creating a compliant presentation...")
    good_pres = ppt.create_presentation()

    ppt.add_slide(good_pres, layout="Title Slide")
    ppt.set_title(good_pres, 1, "Compliant Presentation")
    ppt.set_subtitle(good_pres, 1, "Following all the rules")

    ppt.add_slide(good_pres, layout="Title and Content")
    ppt.set_title(good_pres, 2, "Key Points")
    ppt.set_body(good_pres, 2, [
        "Clear and concise",
        "Easy to read",
        "Well organized",
    ])

    good_issues = validate(good_pres, rules)
    if not good_issues:
        print("   Compliant presentation has no issues!")

    # Save both
    ppt.save_presentation(pres, "with_violations.pptx")
    ppt.save_presentation(good_pres, "compliant.pptx")
    print("\n[5] Saved both presentations")

    print("\nValidation example complete!")


if __name__ == "__main__":
    main()
