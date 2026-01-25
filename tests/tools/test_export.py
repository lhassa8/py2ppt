"""Tests for export functionality."""

import tempfile
from pathlib import Path

import pytest

import py2ppt as ppt


class TestCheckExportDependencies:
    """Tests for check_export_dependencies function."""

    def test_returns_dict(self):
        """Test that check_export_dependencies returns a dict."""
        deps = ppt.check_export_dependencies()

        assert isinstance(deps, dict)
        assert "libreoffice" in deps
        assert "pdftoppm" in deps
        assert "imagemagick" in deps

    def test_values_are_bool(self):
        """Test that dependency values are booleans."""
        deps = ppt.check_export_dependencies()

        for key, value in deps.items():
            assert isinstance(value, bool), f"{key} should be bool, got {type(value)}"


class TestExportToPdf:
    """Tests for export_to_pdf function."""

    def test_export_to_pdf_no_libreoffice(self):
        """Test that export_to_pdf raises error without LibreOffice."""
        deps = ppt.check_export_dependencies()

        if deps["libreoffice"]:
            pytest.skip("LibreOffice is installed, skipping missing dependency test")

        pres = ppt.create_presentation()
        ppt.add_slide(pres, layout="Title Slide")
        ppt.set_title(pres, 1, "Test Export")

        with tempfile.NamedTemporaryFile(suffix=".pdf", delete=False) as f:
            output_path = f.name

        try:
            with pytest.raises(RuntimeError, match="LibreOffice is required"):
                ppt.export_to_pdf(pres, output_path)
        finally:
            Path(output_path).unlink(missing_ok=True)

    def test_export_to_pdf_with_libreoffice(self):
        """Test PDF export when LibreOffice is available."""
        deps = ppt.check_export_dependencies()

        if not deps["libreoffice"]:
            pytest.skip("LibreOffice not installed")

        pres = ppt.create_presentation()
        ppt.add_slide(pres, layout="Title Slide")
        ppt.set_title(pres, 1, "Test Export")

        with tempfile.TemporaryDirectory() as temp_dir:
            output_path = Path(temp_dir) / "output.pdf"

            result = ppt.export_to_pdf(pres, output_path)

            assert result == output_path
            assert output_path.exists()
            assert output_path.stat().st_size > 0


class TestExportSlideToImage:
    """Tests for export_slide_to_image function."""

    def test_export_slide_no_dependencies(self):
        """Test that export_slide_to_image raises error without dependencies."""
        deps = ppt.check_export_dependencies()

        if deps["libreoffice"]:
            pytest.skip("LibreOffice is installed, skipping missing dependency test")

        pres = ppt.create_presentation()
        ppt.add_slide(pres, layout="Title Slide")

        with tempfile.NamedTemporaryFile(suffix=".png", delete=False) as f:
            output_path = f.name

        try:
            with pytest.raises(RuntimeError, match="LibreOffice is required"):
                ppt.export_slide_to_image(pres, 1, output_path)
        finally:
            Path(output_path).unlink(missing_ok=True)

    def test_export_slide_with_dependencies(self):
        """Test slide image export when dependencies are available."""
        deps = ppt.check_export_dependencies()

        if not deps["libreoffice"]:
            pytest.skip("LibreOffice not installed")

        if not deps["pdftoppm"] and not deps["imagemagick"]:
            pytest.skip("Neither pdftoppm nor ImageMagick installed")

        pres = ppt.create_presentation()
        ppt.add_slide(pres, layout="Title Slide")
        ppt.set_title(pres, 1, "Test Slide")

        with tempfile.TemporaryDirectory() as temp_dir:
            output_path = Path(temp_dir) / "slide.png"

            result = ppt.export_slide_to_image(pres, 1, output_path)

            assert result == output_path
            assert output_path.exists()
            assert output_path.stat().st_size > 0


class TestExportAllSlides:
    """Tests for export_all_slides function."""

    def test_export_all_slides_empty_presentation(self):
        """Test exporting slides from empty presentation."""
        deps = ppt.check_export_dependencies()

        if not deps["libreoffice"]:
            pytest.skip("LibreOffice not installed")

        pres = ppt.create_presentation()
        # No slides added

        with tempfile.TemporaryDirectory() as temp_dir:
            result = ppt.export_all_slides(pres, temp_dir)

            assert isinstance(result, list)
            assert len(result) == 0

    def test_export_all_slides_no_dependencies(self):
        """Test that export_all_slides raises error without dependencies."""
        deps = ppt.check_export_dependencies()

        if deps["libreoffice"]:
            pytest.skip("LibreOffice is installed, skipping missing dependency test")

        pres = ppt.create_presentation()
        ppt.add_slide(pres, layout="Title Slide")

        with tempfile.TemporaryDirectory() as temp_dir:
            with pytest.raises(RuntimeError, match="LibreOffice is required"):
                ppt.export_all_slides(pres, temp_dir)

    def test_export_all_slides_with_dependencies(self):
        """Test exporting all slides when dependencies are available."""
        deps = ppt.check_export_dependencies()

        if not deps["libreoffice"]:
            pytest.skip("LibreOffice not installed")

        if not deps["pdftoppm"] and not deps["imagemagick"]:
            pytest.skip("Neither pdftoppm nor ImageMagick installed")

        pres = ppt.create_presentation()
        ppt.add_slide(pres, layout="Title Slide")
        ppt.add_slide(pres, layout="Blank")
        ppt.set_title(pres, 1, "Slide 1")

        with tempfile.TemporaryDirectory() as temp_dir:
            result = ppt.export_all_slides(pres, temp_dir)

            assert isinstance(result, list)
            assert len(result) == 2

            for path in result:
                assert path.exists()
                assert path.stat().st_size > 0


class TestExportFormats:
    """Tests for different export formats."""

    def test_export_jpg_format(self):
        """Test exporting as JPG format."""
        deps = ppt.check_export_dependencies()

        if not deps["libreoffice"] or not (deps["pdftoppm"] or deps["imagemagick"]):
            pytest.skip("Required dependencies not installed")

        pres = ppt.create_presentation()
        ppt.add_slide(pres, layout="Title Slide")

        with tempfile.TemporaryDirectory() as temp_dir:
            output_path = Path(temp_dir) / "slide.jpg"

            result = ppt.export_slide_to_image(pres, 1, output_path, format="jpg")

            assert result == output_path
            assert output_path.exists()

    def test_export_custom_prefix(self):
        """Test exporting with custom prefix."""
        deps = ppt.check_export_dependencies()

        if not deps["libreoffice"] or not (deps["pdftoppm"] or deps["imagemagick"]):
            pytest.skip("Required dependencies not installed")

        pres = ppt.create_presentation()
        ppt.add_slide(pres, layout="Title Slide")

        with tempfile.TemporaryDirectory() as temp_dir:
            result = ppt.export_all_slides(pres, temp_dir, prefix="my_slide")

            assert len(result) == 1
            assert "my_slide" in result[0].name
