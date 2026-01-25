"""3D effects XML handling for shapes.

This module provides dataclasses and utilities for creating 3D rotation,
depth, bevel, and lighting effects on shapes.
"""

from __future__ import annotations

from dataclasses import dataclass, field

from lxml import etree

from .ns import qn


@dataclass
class Rotation3D:
    """3D rotation settings.

    Attributes:
        lat: Latitude rotation (x-axis) in degrees.
        lon: Longitude rotation (y-axis) in degrees.
        rev: Revolution rotation (z-axis) in degrees.
    """

    lat: int = 0
    lon: int = 0
    rev: int = 0


@dataclass
class Bevel:
    """Bevel effect settings.

    Attributes:
        type: Bevel type preset.
        width: Bevel width in EMUs.
        height: Bevel height in EMUs.
    """

    type: str = "circle"
    width: int = 76200  # 6pt in EMUs
    height: int = 76200

    # Available bevel types
    TYPES: list[str] = field(
        default_factory=lambda: [
            "angle",
            "artDeco",
            "circle",
            "convex",
            "coolSlant",
            "cross",
            "divot",
            "hardEdge",
            "relaxedInset",
            "riblet",
            "slope",
            "softRound",
        ],
        repr=False,
    )


@dataclass
class Effect3D:
    """Complete 3D effect configuration.

    Attributes:
        rotation: 3D rotation settings.
        depth: Extrusion depth in EMUs.
        depth_color: Extrusion color (hex or None for auto).
        bevel_top: Top bevel settings.
        bevel_bottom: Bottom bevel settings.
        preset: Camera preset name.
        contour_width: Contour line width in EMUs.
        contour_color: Contour line color (hex).
    """

    rotation: Rotation3D | None = None
    depth: int = 0
    depth_color: str | None = None
    bevel_top: Bevel | None = None
    bevel_bottom: Bevel | None = None
    preset: str | None = None
    contour_width: int = 0
    contour_color: str | None = None


# 3D camera presets
CAMERA_PRESETS = [
    # Parallel (orthographic) presets
    "legacyObliqueTopLeft",
    "legacyObliqueTop",
    "legacyObliqueTopRight",
    "legacyObliqueLeft",
    "legacyObliqueFront",
    "legacyObliqueRight",
    "legacyObliqueBottomLeft",
    "legacyObliqueBottom",
    "legacyObliqueBottomRight",
    "legacyPerspectiveTopLeft",
    "legacyPerspectiveTop",
    "legacyPerspectiveTopRight",
    "legacyPerspectiveLeft",
    "legacyPerspectiveFront",
    "legacyPerspectiveRight",
    "legacyPerspectiveBottomLeft",
    "legacyPerspectiveBottom",
    "legacyPerspectiveBottomRight",
    # Orthographic presets
    "orthographicFront",
    # Isometric presets
    "isometricTopUp",
    "isometricTopDown",
    "isometricBottomUp",
    "isometricBottomDown",
    "isometricLeftUp",
    "isometricLeftDown",
    "isometricRightUp",
    "isometricRightDown",
    "isometricOffAxis1Left",
    "isometricOffAxis1Right",
    "isometricOffAxis1Top",
    "isometricOffAxis2Left",
    "isometricOffAxis2Right",
    "isometricOffAxis2Top",
    "isometricOffAxis3Left",
    "isometricOffAxis3Right",
    "isometricOffAxis3Bottom",
    "isometricOffAxis4Left",
    "isometricOffAxis4Right",
    "isometricOffAxis4Bottom",
    # Oblique presets
    "obliqueTopLeft",
    "obliqueTop",
    "obliqueTopRight",
    "obliqueLeft",
    "obliqueRight",
    "obliqueBottomLeft",
    "obliqueBottom",
    "obliqueBottomRight",
    # Perspective presets
    "perspectiveFront",
    "perspectiveLeft",
    "perspectiveRight",
    "perspectiveAbove",
    "perspectiveBelow",
    "perspectiveAboveLeftFacing",
    "perspectiveAboveRightFacing",
    "perspectiveContrastingLeftFacing",
    "perspectiveContrastingRightFacing",
    "perspectiveHeroicLeftFacing",
    "perspectiveHeroicRightFacing",
    "perspectiveHeroicExtremeLeftFacing",
    "perspectiveHeroicExtremeRightFacing",
    "perspectiveRelaxed",
    "perspectiveRelaxedModerately",
]


def create_scene3d(preset: str = "orthographicFront") -> etree._Element:
    """Create a scene3d element with camera preset.

    Args:
        preset: Camera preset name.

    Returns:
        The a:scene3d element.
    """
    scene3d = etree.Element(qn("a:scene3d"))

    camera = etree.SubElement(scene3d, qn("a:camera"))
    camera.set("prst", preset)

    light_rig = etree.SubElement(scene3d, qn("a:lightRig"))
    light_rig.set("rig", "threePt")
    light_rig.set("dir", "t")

    return scene3d


def create_sp3d(
    depth: int = 0,
    depth_color: str | None = None,
    bevel_top: Bevel | None = None,
    bevel_bottom: Bevel | None = None,
    contour_width: int = 0,
    contour_color: str | None = None,
) -> etree._Element:
    """Create a sp3d element with extrusion and bevel.

    Args:
        depth: Extrusion depth in EMUs.
        depth_color: Extrusion color (hex) or None for auto.
        bevel_top: Top bevel settings.
        bevel_bottom: Bottom bevel settings.
        contour_width: Contour width in EMUs.
        contour_color: Contour color (hex).

    Returns:
        The a:sp3d element.
    """
    sp3d = etree.Element(qn("a:sp3d"))

    if depth > 0:
        sp3d.set("extrusionH", str(depth))

    if contour_width > 0:
        sp3d.set("contourW", str(contour_width))

    # Bevel top
    if bevel_top is not None:
        bevel_t = etree.SubElement(sp3d, qn("a:bevelT"))
        bevel_t.set("prst", bevel_top.type)
        bevel_t.set("w", str(bevel_top.width))
        bevel_t.set("h", str(bevel_top.height))

    # Bevel bottom
    if bevel_bottom is not None:
        bevel_b = etree.SubElement(sp3d, qn("a:bevelB"))
        bevel_b.set("prst", bevel_bottom.type)
        bevel_b.set("w", str(bevel_bottom.width))
        bevel_b.set("h", str(bevel_bottom.height))

    # Extrusion color
    if depth > 0 and depth_color is not None:
        extrusion_clr = etree.SubElement(sp3d, qn("a:extrusionClr"))
        srgb = etree.SubElement(extrusion_clr, qn("a:srgbClr"))
        srgb.set("val", depth_color.lstrip("#").upper())

    # Contour color
    if contour_width > 0 and contour_color is not None:
        contour_clr = etree.SubElement(sp3d, qn("a:contourClr"))
        srgb = etree.SubElement(contour_clr, qn("a:srgbClr"))
        srgb.set("val", contour_color.lstrip("#").upper())

    return sp3d


def create_rotation_element(rotation: Rotation3D) -> etree._Element:
    """Create a rot element for 3D rotation.

    Args:
        rotation: Rotation settings.

    Returns:
        The a:rot element.
    """
    rot = etree.Element(qn("a:rot"))

    # Convert degrees to 60000ths of a degree
    rot.set("lat", str(rotation.lat * 60000))
    rot.set("lon", str(rotation.lon * 60000))
    rot.set("rev", str(rotation.rev * 60000))

    return rot
