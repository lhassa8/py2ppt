"""Export tool functions.

Functions for exporting presentations to PDF and images.
Requires external tools (LibreOffice) for PDF export.
"""

from __future__ import annotations

import os
import shutil
import subprocess
import tempfile
from pathlib import Path
from typing import Literal

from ..core.presentation import Presentation


def _find_libreoffice() -> str | None:
    """Find LibreOffice executable path."""
    # Common paths for LibreOffice
    common_paths = [
        # macOS
        "/Applications/LibreOffice.app/Contents/MacOS/soffice",
        # Linux
        "/usr/bin/libreoffice",
        "/usr/bin/soffice",
        "/usr/local/bin/libreoffice",
        "/usr/local/bin/soffice",
        # Windows
        r"C:\Program Files\LibreOffice\program\soffice.exe",
        r"C:\Program Files (x86)\LibreOffice\program\soffice.exe",
    ]

    # Check common paths
    for path in common_paths:
        if os.path.isfile(path):
            return path

    # Try to find in PATH
    soffice = shutil.which("soffice")
    if soffice:
        return soffice

    libreoffice = shutil.which("libreoffice")
    if libreoffice:
        return libreoffice

    return None


def export_to_pdf(
    presentation: Presentation,
    output_path: str | Path,
) -> Path:
    """Export presentation to PDF.

    Requires LibreOffice to be installed on the system.

    Args:
        presentation: The presentation to export
        output_path: Path for the output PDF file

    Returns:
        Path to the created PDF file

    Raises:
        RuntimeError: If LibreOffice is not installed

    Example:
        >>> pres = ppt.create_presentation()
        >>> ppt.add_slide(pres, layout="Title Slide")
        >>> ppt.set_title(pres, 1, "Hello World")
        >>> pdf_path = ppt.export_to_pdf(pres, "output.pdf")
    """
    soffice = _find_libreoffice()
    if soffice is None:
        raise RuntimeError(
            "LibreOffice is required for PDF export. "
            "Please install LibreOffice: https://www.libreoffice.org/download/"
        )

    output_path = Path(output_path)
    output_dir = output_path.parent
    output_dir.mkdir(parents=True, exist_ok=True)

    # Save presentation to a temp file first
    with tempfile.NamedTemporaryFile(suffix=".pptx", delete=False) as f:
        temp_pptx = Path(f.name)

    try:
        # Save the presentation
        presentation.save(str(temp_pptx))

        # Convert using LibreOffice
        cmd = [
            soffice,
            "--headless",
            "--convert-to", "pdf",
            "--outdir", str(output_dir),
            str(temp_pptx),
        ]

        result = subprocess.run(
            cmd,
            capture_output=True,
            text=True,
            timeout=120,  # 2 minute timeout
        )

        if result.returncode != 0:
            raise RuntimeError(
                f"LibreOffice conversion failed: {result.stderr}"
            )

        # LibreOffice creates file with same name but .pdf extension
        converted_pdf = output_dir / (temp_pptx.stem + ".pdf")

        # Rename to desired output path if different
        if converted_pdf != output_path:
            if output_path.exists():
                output_path.unlink()
            converted_pdf.rename(output_path)

        return output_path

    finally:
        # Clean up temp file
        if temp_pptx.exists():
            temp_pptx.unlink()


def export_slide_to_image(
    presentation: Presentation,
    slide_number: int,
    output_path: str | Path,
    *,
    format: Literal["png", "jpg", "jpeg", "gif"] = "png",
    width: int | None = None,
    height: int | None = None,
) -> Path:
    """Export a single slide to an image.

    Requires LibreOffice or another image conversion tool.
    If width/height specified, exports at that resolution.

    Args:
        presentation: The presentation to export
        slide_number: The slide number (1-indexed)
        output_path: Path for the output image file
        format: Image format ("png", "jpg", "jpeg", "gif")
        width: Output image width in pixels (optional)
        height: Output image height in pixels (optional)

    Returns:
        Path to the created image file

    Raises:
        RuntimeError: If conversion tools are not available

    Example:
        >>> pres = ppt.open_presentation("slides.pptx")
        >>> img_path = ppt.export_slide_to_image(pres, 1, "slide1.png")
    """
    soffice = _find_libreoffice()
    if soffice is None:
        raise RuntimeError(
            "LibreOffice is required for image export. "
            "Please install LibreOffice: https://www.libreoffice.org/download/"
        )

    output_path = Path(output_path)
    output_dir = output_path.parent
    output_dir.mkdir(parents=True, exist_ok=True)

    # Normalize format
    if format == "jpeg":
        format = "jpg"

    # First export to PDF, then convert to image
    # (LibreOffice doesn't support direct per-slide image export easily)

    with tempfile.TemporaryDirectory() as temp_dir:
        temp_dir_path = Path(temp_dir)

        # Save the presentation
        temp_pptx = temp_dir_path / "presentation.pptx"
        presentation.save(str(temp_pptx))

        # Convert to PDF first
        cmd = [
            soffice,
            "--headless",
            "--convert-to", "pdf",
            "--outdir", str(temp_dir_path),
            str(temp_pptx),
        ]

        result = subprocess.run(
            cmd,
            capture_output=True,
            text=True,
            timeout=120,
        )

        if result.returncode != 0:
            raise RuntimeError(
                f"LibreOffice PDF conversion failed: {result.stderr}"
            )

        temp_pdf = temp_dir_path / "presentation.pdf"

        if not temp_pdf.exists():
            raise RuntimeError("PDF conversion did not produce output")

        # Try to use pdftoppm (poppler-utils) or ImageMagick for PDF to image
        # Check for pdftoppm
        pdftoppm = shutil.which("pdftoppm")
        if pdftoppm:
            # Use pdftoppm for PDF to image conversion
            # -f and -l specify first and last page
            if format == "png":
                fmt_flag = "-png"
            elif format in ("jpg", "jpeg"):
                fmt_flag = "-jpeg"
            else:
                fmt_flag = "-png"
                format = "png"

            img_prefix = temp_dir_path / "slide"

            cmd = [
                pdftoppm,
                fmt_flag,
                "-f", str(slide_number),
                "-l", str(slide_number),
                str(temp_pdf),
                str(img_prefix),
            ]

            if width:
                cmd.extend(["-scale-to-x", str(width)])
            if height:
                cmd.extend(["-scale-to-y", str(height)])

            result = subprocess.run(
                cmd,
                capture_output=True,
                text=True,
                timeout=60,
            )

            if result.returncode != 0:
                raise RuntimeError(
                    f"pdftoppm conversion failed: {result.stderr}"
                )

            # Find the output file (pdftoppm adds page number)
            # Format: slide-01.png or slide-1.png depending on total pages
            possible_names = [
                f"slide-{slide_number}.{format}",
                f"slide-{slide_number:02d}.{format}",
                f"slide-{slide_number:03d}.{format}",
            ]

            img_file = None
            for name in possible_names:
                candidate = temp_dir_path / name
                if candidate.exists():
                    img_file = candidate
                    break

            if img_file is None:
                # Try glob pattern
                matches = list(temp_dir_path.glob(f"slide-*.{format}"))
                if matches:
                    img_file = matches[0]

            if img_file and img_file.exists():
                shutil.copy(img_file, output_path)
                return output_path

        # Try ImageMagick convert
        convert = shutil.which("convert")
        if convert:
            # ImageMagick uses 0-based page indexing
            page_spec = f"[{slide_number - 1}]"

            cmd = [
                convert,
                "-density", "150",
                str(temp_pdf) + page_spec,
                str(output_path),
            ]

            if width and height:
                cmd.insert(-1, "-resize")
                cmd.insert(-1, f"{width}x{height}")

            result = subprocess.run(
                cmd,
                capture_output=True,
                text=True,
                timeout=60,
            )

            if result.returncode == 0 and output_path.exists():
                return output_path

        raise RuntimeError(
            "Could not convert slide to image. "
            "Please install pdftoppm (poppler-utils) or ImageMagick."
        )


def export_all_slides(
    presentation: Presentation,
    output_dir: str | Path,
    *,
    format: Literal["png", "jpg", "jpeg", "gif"] = "png",
    prefix: str = "slide",
    width: int | None = None,
    height: int | None = None,
) -> list[Path]:
    """Export all slides to images.

    Each slide is exported as a separate image file.

    Args:
        presentation: The presentation to export
        output_dir: Directory for output images
        format: Image format ("png", "jpg", "jpeg", "gif")
        prefix: Filename prefix (default "slide")
        width: Output image width in pixels (optional)
        height: Output image height in pixels (optional)

    Returns:
        List of paths to created image files

    Raises:
        RuntimeError: If conversion tools are not available

    Example:
        >>> pres = ppt.open_presentation("slides.pptx")
        >>> images = ppt.export_all_slides(pres, "images/", format="png")
        >>> for img in images:
        ...     print(f"Created: {img}")
    """
    output_dir = Path(output_dir)
    output_dir.mkdir(parents=True, exist_ok=True)

    # Get slide count
    from ..tools.inspection import get_slide_count
    num_slides = get_slide_count(presentation)

    if num_slides == 0:
        return []

    # Normalize format
    if format == "jpeg":
        format = "jpg"

    # Check for LibreOffice
    soffice = _find_libreoffice()
    if soffice is None:
        raise RuntimeError(
            "LibreOffice is required for image export. "
            "Please install LibreOffice: https://www.libreoffice.org/download/"
        )

    result_paths = []

    # Try batch export via pdftoppm for efficiency
    pdftoppm = shutil.which("pdftoppm")
    if pdftoppm:
        with tempfile.TemporaryDirectory() as temp_dir:
            temp_dir_path = Path(temp_dir)

            # Save and convert to PDF
            temp_pptx = temp_dir_path / "presentation.pptx"
            presentation.save(str(temp_pptx))

            cmd = [
                soffice,
                "--headless",
                "--convert-to", "pdf",
                "--outdir", str(temp_dir_path),
                str(temp_pptx),
            ]

            result = subprocess.run(
                cmd,
                capture_output=True,
                text=True,
                timeout=120,
            )

            if result.returncode != 0:
                raise RuntimeError(
                    f"LibreOffice PDF conversion failed: {result.stderr}"
                )

            temp_pdf = temp_dir_path / "presentation.pdf"

            if not temp_pdf.exists():
                raise RuntimeError("PDF conversion did not produce output")

            # Convert all pages at once
            if format == "png":
                fmt_flag = "-png"
            elif format in ("jpg", "jpeg"):
                fmt_flag = "-jpeg"
            else:
                fmt_flag = "-png"
                format = "png"

            img_prefix = temp_dir_path / "page"

            cmd = [
                pdftoppm,
                fmt_flag,
                str(temp_pdf),
                str(img_prefix),
            ]

            if width:
                cmd.extend(["-scale-to-x", str(width)])
            if height:
                cmd.extend(["-scale-to-y", str(height)])

            result = subprocess.run(
                cmd,
                capture_output=True,
                text=True,
                timeout=300,  # 5 minute timeout for many slides
            )

            if result.returncode != 0:
                raise RuntimeError(
                    f"pdftoppm conversion failed: {result.stderr}"
                )

            # Find and rename all output files
            for slide_num in range(1, num_slides + 1):
                # Try different naming conventions
                possible_names = [
                    f"page-{slide_num}.{format}",
                    f"page-{slide_num:02d}.{format}",
                    f"page-{slide_num:03d}.{format}",
                ]

                img_file = None
                for name in possible_names:
                    candidate = temp_dir_path / name
                    if candidate.exists():
                        img_file = candidate
                        break

                if img_file and img_file.exists():
                    dest_path = output_dir / f"{prefix}_{slide_num:03d}.{format}"
                    shutil.copy(img_file, dest_path)
                    result_paths.append(dest_path)

            return result_paths

    # Fall back to per-slide export
    for slide_num in range(1, num_slides + 1):
        output_path = output_dir / f"{prefix}_{slide_num:03d}.{format}"
        export_slide_to_image(
            presentation,
            slide_num,
            output_path,
            format=format,
            width=width,
            height=height,
        )
        result_paths.append(output_path)

    return result_paths


def check_export_dependencies() -> dict[str, bool]:
    """Check which export dependencies are available.

    Returns:
        Dict with dependency name -> availability status

    Example:
        >>> deps = ppt.check_export_dependencies()
        >>> if deps["libreoffice"]:
        ...     print("PDF export is available")
    """
    return {
        "libreoffice": _find_libreoffice() is not None,
        "pdftoppm": shutil.which("pdftoppm") is not None,
        "imagemagick": shutil.which("convert") is not None,
    }
