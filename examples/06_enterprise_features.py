#!/usr/bin/env python3
"""Example: Enterprise PowerPoint Features.

This example demonstrates the advanced features of py2ppt including:
- Charts (column, pie, line)
- Diagrams (process, cycle, hierarchy)
- Shapes with gradients and connectors
- Animations and transitions
- Theme customization
- Style validation

Run this script to generate a comprehensive presentation showcasing
all enterprise features.
"""

import py2ppt as ppt


def create_enterprise_presentation():
    """Create a presentation demonstrating enterprise features."""

    # Create presentation
    pres = ppt.create_presentation()

    # =========================================================================
    # Slide 1: Title Slide with Transition
    # =========================================================================
    ppt.add_slide(pres, layout="Title Slide")
    ppt.set_title(pres, 1, "Enterprise PowerPoint Features")
    ppt.set_subtitle(pres, 1, "Charts, Diagrams, Animations & More")

    # Add a fade transition
    ppt.set_slide_transition(pres, 1, transition="fade", duration=1000)

    # =========================================================================
    # Slide 2: Column Chart
    # =========================================================================
    ppt.add_slide(pres, layout="Title and Content")
    ppt.set_title(pres, 2, "Quarterly Revenue")

    ppt.add_chart(
        pres,
        2,
        "column",
        categories=["Q1", "Q2", "Q3", "Q4"],
        series=[
            {"name": "2023", "values": [100, 120, 140, 160]},
            {"name": "2024", "values": [110, 135, 155, 180]},
        ],
        title="Revenue by Quarter",
        legend="right",
        left="1in",
        top="1.5in",
        width="8in",
        height="4.5in",
    )

    ppt.set_slide_transition(pres, 2, transition="push", direction="left")

    # =========================================================================
    # Slide 3: Pie Chart
    # =========================================================================
    ppt.add_slide(pres, layout="Title and Content")
    ppt.set_title(pres, 3, "Market Share Distribution")

    ppt.add_chart(
        pres,
        3,
        "pie",
        categories=["North America", "Europe", "Asia Pacific", "Rest of World"],
        series=[{"name": "Market Share", "values": [35, 28, 25, 12]}],
        data_labels=True,
        left="2in",
        top="1.5in",
        width="6in",
        height="5in",
    )

    ppt.set_slide_transition(pres, 3, transition="wipe")

    # =========================================================================
    # Slide 4: Line Chart
    # =========================================================================
    ppt.add_slide(pres, layout="Title and Content")
    ppt.set_title(pres, 4, "Growth Trends")

    ppt.add_chart(
        pres,
        4,
        "line",
        categories=["Jan", "Feb", "Mar", "Apr", "May", "Jun"],
        series=[
            {"name": "Actual", "values": [10, 15, 13, 18, 22, 28]},
            {"name": "Target", "values": [12, 14, 16, 18, 20, 22]},
        ],
        markers=True,
        left="1in",
        top="1.5in",
        width="8in",
        height="4.5in",
    )

    # =========================================================================
    # Slide 5: Process Diagram
    # =========================================================================
    ppt.add_slide(pres, layout="Title and Content")
    ppt.set_title(pres, 5, "Development Process")

    ppt.add_diagram(
        pres,
        5,
        "process",
        ["Research", "Design", "Develop", "Test", "Deploy"],
        left="0.5in",
        top="1.8in",
        width="9in",
        height="2in",
    )

    ppt.set_slide_transition(pres, 5, transition="fade")

    # =========================================================================
    # Slide 6: Cycle Diagram
    # =========================================================================
    ppt.add_slide(pres, layout="Title and Content")
    ppt.set_title(pres, 6, "Continuous Improvement")

    ppt.add_diagram(
        pres,
        6,
        "cycle",
        ["Plan", "Do", "Check", "Act"],
        left="2in",
        top="1.5in",
        width="6in",
        height="5in",
    )

    # =========================================================================
    # Slide 7: Hierarchy Diagram (Org Chart)
    # =========================================================================
    ppt.add_slide(pres, layout="Title and Content")
    ppt.set_title(pres, 7, "Organization Structure")

    ppt.add_diagram(
        pres,
        7,
        "hierarchy",
        [
            {
                "text": "CEO",
                "children": [
                    {
                        "text": "CTO",
                        "children": [
                            {"text": "Engineering"},
                            {"text": "Product"},
                        ],
                    },
                    {"text": "CFO"},
                    {"text": "COO"},
                ],
            }
        ],
        left="0.5in",
        top="1.5in",
        width="9in",
        height="5in",
    )

    # =========================================================================
    # Slide 8: Shapes and Connectors
    # =========================================================================
    ppt.add_slide(pres, layout="Blank")

    # Add title manually
    ppt.add_shape(
        pres,
        8,
        "rectangle",
        left="0.5in",
        top="0.3in",
        width="9in",
        height="0.6in",
        text="Custom Shapes & Connectors",
        fill="none",
        outline=False,
    )

    # Add shapes
    shape1 = ppt.add_shape(
        pres,
        8,
        "rounded_rectangle",
        left="1in",
        top="2in",
        width="2in",
        height="1.5in",
        text="Step 1",
        fill="accent1",
    )

    shape2 = ppt.add_shape(
        pres,
        8,
        "rounded_rectangle",
        left="4in",
        top="2in",
        width="2in",
        height="1.5in",
        text="Step 2",
        fill="accent2",
    )

    shape3 = ppt.add_shape(
        pres,
        8,
        "rounded_rectangle",
        left="7in",
        top="2in",
        width="2in",
        height="1.5in",
        text="Step 3",
        fill="accent3",
    )

    # Add connectors
    ppt.add_connector(
        pres,
        8,
        start_x="3in",
        start_y="2.75in",
        end_x="4in",
        end_y="2.75in",
        end_arrow="triangle",
    )

    ppt.add_connector(
        pres,
        8,
        start_x="6in",
        start_y="2.75in",
        end_x="7in",
        end_y="2.75in",
        end_arrow="triangle",
    )

    # Add a gradient shape
    ppt.add_shape(
        pres,
        8,
        "ellipse",
        left="3.5in",
        top="4.5in",
        width="3in",
        height="2in",
        text="Gradient Fill",
        fill={"type": "gradient", "colors": ["#FF6B6B", "#4ECDC4"]},
    )

    # =========================================================================
    # Slide 9: Pyramid Diagram
    # =========================================================================
    ppt.add_slide(pres, layout="Title and Content")
    ppt.set_title(pres, 9, "Priority Levels")

    ppt.add_diagram(
        pres,
        9,
        "pyramid",
        ["Critical", "High", "Medium", "Low"],
        left="2in",
        top="1.5in",
        width="6in",
        height="5in",
    )

    # =========================================================================
    # Slide 10: Venn Diagram
    # =========================================================================
    ppt.add_slide(pres, layout="Title and Content")
    ppt.set_title(pres, 10, "Finding Your Ikigai")

    ppt.add_diagram(
        pres,
        10,
        "venn",
        ["What you love", "What you're good at", "What the world needs"],
        left="1.5in",
        top="1.5in",
        width="7in",
        height="5in",
    )

    # =========================================================================
    # Slide 11: Styled Table
    # =========================================================================
    ppt.add_slide(pres, layout="Title and Content")
    ppt.set_title(pres, 11, "Sales Report")

    ppt.add_table(
        pres,
        11,
        data=[
            ["Region", "Q1", "Q2", "Q3", "Q4", "Total"],
            ["North", 150, 180, 200, 220, 750],
            ["South", 120, 140, 160, 180, 600],
            ["East", 100, 120, 130, 150, 500],
            ["West", 90, 100, 120, 140, 450],
        ],
        left="1in",
        top="1.5in",
        width="8in",
        height="4in",
        header_row=True,
        banded_rows=True,
        header_background="#003366",
        header_text_color="#FFFFFF",
    )

    # Style specific cells
    ppt.style_table_cell(pres, 11, 0, row=1, col=5, bold=True, background="#E6F3FF")
    ppt.style_table_cell(pres, 11, 0, row=2, col=5, bold=True, background="#E6F3FF")
    ppt.style_table_cell(pres, 11, 0, row=3, col=5, bold=True, background="#E6F3FF")
    ppt.style_table_cell(pres, 11, 0, row=4, col=5, bold=True, background="#E6F3FF")

    # =========================================================================
    # Slide 12: Thank You with Animation
    # =========================================================================
    ppt.add_slide(pres, layout="Title Slide")
    ppt.set_title(pres, 12, "Thank You!")
    ppt.set_subtitle(pres, 12, "Questions?")

    ppt.set_slide_transition(pres, 12, transition="fade", duration=1500)

    # =========================================================================
    # Apply Theme Customization
    # =========================================================================
    # Customize theme colors for brand consistency
    ppt.apply_theme_colors(
        pres,
        {
            "accent1": "#0066CC",  # Primary blue
            "accent2": "#00A651",  # Secondary green
            "accent3": "#FF6B00",  # Accent orange
        },
    )

    # Save the presentation
    ppt.save_presentation(pres, "enterprise_features_demo.pptx")
    print("Created: enterprise_features_demo.pptx")

    # =========================================================================
    # Validation Example
    # =========================================================================
    from py2ppt.validation import (
        corporate_style_guide,
        get_validation_summary,
        validate,
    )

    # Validate against corporate style guide
    guide = corporate_style_guide()
    issues = validate(pres, guide)
    summary = get_validation_summary(issues)

    print(f"\nValidation Results:")
    print(f"  Total issues: {summary['total']}")
    print(f"  Errors: {summary['error_count']}")
    print(f"  Warnings: {summary['warning_count']}")
    print(f"  Info: {summary['info_count']}")

    if issues:
        print("\nTop issues:")
        for issue in issues[:5]:
            slide_info = f"Slide {issue.slide}" if issue.slide else "Presentation"
            print(f"  [{issue.severity}] {slide_info}: {issue.message}")


if __name__ == "__main__":
    create_enterprise_presentation()
