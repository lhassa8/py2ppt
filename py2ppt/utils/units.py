"""Unit conversions for PowerPoint.

PowerPoint uses EMUs (English Metric Units) internally.
914400 EMUs = 1 inch
"""

from __future__ import annotations

import re

# EMUs per inch
EMU_PER_INCH = 914400
EMU_PER_CM = 360000
EMU_PER_MM = 36000
EMU_PER_PT = 12700  # 1 point = 1/72 inch


class EMU(int):
    """English Metric Unit value.

    This is the native unit used in PowerPoint XML.
    914400 EMUs = 1 inch
    """

    def __new__(cls, value: int) -> EMU:
        return super().__new__(cls, int(value))

    @property
    def inches(self) -> float:
        """Convert to inches."""
        return self / EMU_PER_INCH

    @property
    def cm(self) -> float:
        """Convert to centimeters."""
        return self / EMU_PER_CM

    @property
    def mm(self) -> float:
        """Convert to millimeters."""
        return self / EMU_PER_MM

    @property
    def pt(self) -> float:
        """Convert to points."""
        return self / EMU_PER_PT

    def __repr__(self) -> str:
        return f"EMU({int(self)})"


def Emu(value: int) -> EMU:
    """Create an EMU value. Alias for EMU()."""
    return EMU(value)


def Inches(value: float) -> EMU:
    """Convert inches to EMUs."""
    return EMU(int(value * EMU_PER_INCH))


def Cm(value: float) -> EMU:
    """Convert centimeters to EMUs."""
    return EMU(int(value * EMU_PER_CM))


def Mm(value: float) -> EMU:
    """Convert millimeters to EMUs."""
    return EMU(int(value * EMU_PER_MM))


def Pt(value: float) -> EMU:
    """Convert points to EMUs."""
    return EMU(int(value * EMU_PER_PT))


# Conversion functions
def inches_to_emu(inches: float) -> int:
    """Convert inches to EMUs."""
    return int(inches * EMU_PER_INCH)


def cm_to_emu(cm: float) -> int:
    """Convert centimeters to EMUs."""
    return int(cm * EMU_PER_CM)


def mm_to_emu(mm: float) -> int:
    """Convert millimeters to EMUs."""
    return int(mm * EMU_PER_MM)


def pt_to_emu(pt: float) -> int:
    """Convert points to EMUs."""
    return int(pt * EMU_PER_PT)


def emu_to_inches(emu: int) -> float:
    """Convert EMUs to inches."""
    return emu / EMU_PER_INCH


def emu_to_cm(emu: int) -> float:
    """Convert EMUs to centimeters."""
    return emu / EMU_PER_CM


def emu_to_pt(emu: int) -> float:
    """Convert EMUs to points."""
    return emu / EMU_PER_PT


# Length string parsing
_LENGTH_PATTERN = re.compile(
    r"^\s*(-?\d+(?:\.\d+)?)\s*(in|inch|inches|cm|mm|pt|emu|px)?\s*$",
    re.IGNORECASE,
)


def parse_length(value: str | int | float | EMU) -> EMU:
    """Parse a length value to EMUs.

    Supports:
    - Numbers (treated as EMUs)
    - Strings with units: "1in", "2.5cm", "10mm", "12pt", "914400emu"
    - Strings without units (treated as EMUs)
    - EMU objects (returned as-is)

    Args:
        value: Length value to parse

    Returns:
        EMU value

    Examples:
        parse_length("1in") -> EMU(914400)
        parse_length("2.54cm") -> EMU(914400)
        parse_length("25.4mm") -> EMU(914400)
        parse_length("72pt") -> EMU(914400)
        parse_length(914400) -> EMU(914400)
    """
    if isinstance(value, EMU):
        return value

    if isinstance(value, (int, float)):
        return EMU(int(value))

    if not isinstance(value, str):
        raise ValueError(f"Cannot parse length from {type(value).__name__}")

    match = _LENGTH_PATTERN.match(value)
    if not match:
        raise ValueError(f"Invalid length format: {value}")

    num = float(match.group(1))
    unit = (match.group(2) or "").lower()

    if unit in ("in", "inch", "inches"):
        return Inches(num)
    elif unit == "cm":
        return Cm(num)
    elif unit == "mm":
        return Mm(num)
    elif unit == "pt":
        return Pt(num)
    elif unit == "emu" or not unit:
        return EMU(int(num))
    elif unit == "px":
        # Assume 96 DPI
        return Inches(num / 96)
    else:
        raise ValueError(f"Unknown unit: {unit}")


def format_emu(emu: int, unit: str = "in") -> str:
    """Format an EMU value as a string with units.

    Args:
        emu: EMU value
        unit: Unit to use ("in", "cm", "mm", "pt")

    Returns:
        Formatted string like "1.5in"
    """
    if unit in ("in", "inch", "inches"):
        return f"{emu / EMU_PER_INCH:.2f}in"
    elif unit == "cm":
        return f"{emu / EMU_PER_CM:.2f}cm"
    elif unit == "mm":
        return f"{emu / EMU_PER_MM:.1f}mm"
    elif unit == "pt":
        return f"{emu / EMU_PER_PT:.1f}pt"
    else:
        return str(emu)
