"""Tests for diagram functionality."""

import pytest

import py2ppt as ppt


class TestAddDiagram:
    """Tests for add_diagram function."""

    def test_add_process_diagram(self, tmp_path):
        """Test adding a process flow diagram."""
        pres = ppt.create_presentation()
        ppt.add_slide(pres, layout="Blank")

        ppt.add_diagram(
            pres,
            1,
            "process",
            ["Research", "Design", "Build", "Test", "Launch"],
        )

        path = tmp_path / "process_diagram.pptx"
        ppt.save_presentation(pres, str(path))
        assert path.exists()

    def test_add_cycle_diagram(self, tmp_path):
        """Test adding a cycle diagram."""
        pres = ppt.create_presentation()
        ppt.add_slide(pres, layout="Blank")

        ppt.add_diagram(
            pres,
            1,
            "cycle",
            ["Plan", "Do", "Check", "Act"],
        )

        path = tmp_path / "cycle_diagram.pptx"
        ppt.save_presentation(pres, str(path))
        assert path.exists()

    def test_add_pyramid_diagram(self, tmp_path):
        """Test adding a pyramid diagram."""
        pres = ppt.create_presentation()
        ppt.add_slide(pres, layout="Blank")

        ppt.add_diagram(
            pres,
            1,
            "pyramid",
            ["Leadership", "Management", "Staff", "Contractors"],
        )

        path = tmp_path / "pyramid_diagram.pptx"
        ppt.save_presentation(pres, str(path))
        assert path.exists()

    def test_add_hierarchy_diagram(self, tmp_path):
        """Test adding a hierarchy/org chart diagram."""
        pres = ppt.create_presentation()
        ppt.add_slide(pres, layout="Blank")

        ppt.add_diagram(
            pres,
            1,
            "hierarchy",
            [
                {
                    "text": "CEO",
                    "children": [
                        {"text": "CTO"},
                        {"text": "CFO"},
                        {"text": "COO"},
                    ],
                }
            ],
        )

        path = tmp_path / "hierarchy_diagram.pptx"
        ppt.save_presentation(pres, str(path))
        assert path.exists()

    def test_add_venn_diagram(self, tmp_path):
        """Test adding a Venn diagram."""
        pres = ppt.create_presentation()
        ppt.add_slide(pres, layout="Blank")

        ppt.add_diagram(
            pres,
            1,
            "venn",
            ["Skills", "Passion", "Market Need"],
        )

        path = tmp_path / "venn_diagram.pptx"
        ppt.save_presentation(pres, str(path))
        assert path.exists()

    def test_add_list_diagram(self, tmp_path):
        """Test adding a list diagram."""
        pres = ppt.create_presentation()
        ppt.add_slide(pres, layout="Blank")

        ppt.add_diagram(
            pres,
            1,
            "list",
            ["First item", "Second item", "Third item"],
            title="Key Points",
        )

        path = tmp_path / "list_diagram.pptx"
        ppt.save_presentation(pres, str(path))
        assert path.exists()

    def test_add_radial_diagram(self, tmp_path):
        """Test adding a radial/hub-spoke diagram."""
        pres = ppt.create_presentation()
        ppt.add_slide(pres, layout="Blank")

        ppt.add_diagram(
            pres,
            1,
            "radial",
            ["Core", "Feature 1", "Feature 2", "Feature 3", "Feature 4"],
        )

        path = tmp_path / "radial_diagram.pptx"
        ppt.save_presentation(pres, str(path))
        assert path.exists()

    def test_invalid_diagram_type(self):
        """Test that invalid diagram type raises error."""
        pres = ppt.create_presentation()
        ppt.add_slide(pres, layout="Blank")

        with pytest.raises(ValueError, match="Unknown diagram type"):
            ppt.add_diagram(pres, 1, "invalid_type", ["A", "B", "C"])


class TestGetDiagramTypes:
    """Tests for get_diagram_types function."""

    def test_get_diagram_types(self):
        """Test getting available diagram types."""
        types = ppt.get_diagram_types()
        assert isinstance(types, dict)
        assert "process" in types
        assert "cycle" in types
        assert "hierarchy" in types
        assert "pyramid" in types
        assert "venn" in types
