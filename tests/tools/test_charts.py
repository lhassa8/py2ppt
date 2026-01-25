"""Tests for chart functionality."""


import py2ppt as ppt


class TestAddChart:
    """Tests for add_chart function."""

    def test_add_column_chart(self, tmp_path):
        """Test adding a basic column chart."""
        pres = ppt.create_presentation()
        ppt.add_slide(pres, layout="Title and Content")

        ppt.add_chart(
            pres,
            1,
            "column",
            categories=["Q1", "Q2", "Q3", "Q4"],
            series=[
                {"name": "2023", "values": [100, 120, 140, 160]},
                {"name": "2024", "values": [110, 135, 155, 180]},
            ],
            title="Quarterly Revenue",
        )

        # Save and verify it can be reopened
        path = tmp_path / "chart_test.pptx"
        ppt.save_presentation(pres, str(path))

        reopened = ppt.open_presentation(str(path))
        assert ppt.get_slide_count(reopened) == 1

    def test_add_pie_chart(self, tmp_path):
        """Test adding a pie chart."""
        pres = ppt.create_presentation()
        ppt.add_slide(pres, layout="Title and Content")

        ppt.add_chart(
            pres,
            1,
            "pie",
            categories=["North", "South", "East", "West"],
            series=[{"name": "Sales", "values": [30, 25, 20, 25]}],
            data_labels=True,
        )

        path = tmp_path / "pie_chart.pptx"
        ppt.save_presentation(pres, str(path))
        assert path.exists()

    def test_add_line_chart(self, tmp_path):
        """Test adding a line chart."""
        pres = ppt.create_presentation()
        ppt.add_slide(pres, layout="Title and Content")

        ppt.add_chart(
            pres,
            1,
            "line",
            categories=["Jan", "Feb", "Mar", "Apr", "May"],
            series=[
                {"name": "Actual", "values": [10, 15, 13, 17, 20]},
                {"name": "Target", "values": [12, 14, 16, 18, 20]},
            ],
            markers=True,
        )

        path = tmp_path / "line_chart.pptx"
        ppt.save_presentation(pres, str(path))
        assert path.exists()

    def test_add_bar_chart(self, tmp_path):
        """Test adding a horizontal bar chart."""
        pres = ppt.create_presentation()
        ppt.add_slide(pres, layout="Title and Content")

        ppt.add_chart(
            pres,
            1,
            "bar",
            categories=["Product A", "Product B", "Product C"],
            series=[{"name": "Revenue", "values": [500, 750, 600]}],
        )

        path = tmp_path / "bar_chart.pptx"
        ppt.save_presentation(pres, str(path))
        assert path.exists()

    def test_chart_with_position(self, tmp_path):
        """Test adding a chart with explicit position."""
        pres = ppt.create_presentation()
        ppt.add_slide(pres, layout="Blank")

        ppt.add_chart(
            pres,
            1,
            "column",
            categories=["A", "B", "C"],
            series=[{"name": "Data", "values": [1, 2, 3]}],
            left="1in",
            top="1in",
            width="6in",
            height="4in",
        )

        path = tmp_path / "positioned_chart.pptx"
        ppt.save_presentation(pres, str(path))
        assert path.exists()


class TestUpdateChartData:
    """Tests for update_chart_data function."""

    def test_update_chart_data(self, tmp_path):
        """Test updating chart data."""
        pres = ppt.create_presentation()
        ppt.add_slide(pres, layout="Title and Content")

        # Add initial chart
        ppt.add_chart(
            pres,
            1,
            "column",
            categories=["A", "B", "C"],
            series=[{"name": "Original", "values": [1, 2, 3]}],
        )

        # Update the data
        ppt.update_chart_data(
            pres,
            1,
            chart_index=0,
            categories=["X", "Y", "Z"],
            series=[{"name": "Updated", "values": [10, 20, 30]}],
        )

        path = tmp_path / "updated_chart.pptx"
        ppt.save_presentation(pres, str(path))
        assert path.exists()
