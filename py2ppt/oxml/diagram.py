"""Diagram/SmartArt XML handling.

Provides dataclasses and XML generation for DrawingML diagrams
(SmartArt) in PresentationML format.

Note: SmartArt is complex, requiring multiple XML parts. This module
provides a simplified interface for common diagram types.
"""

from __future__ import annotations

from dataclasses import dataclass, field
from typing import Any

from lxml import etree

from .ns import nsmap, qn

# Diagram type definitions
DIAGRAM_TYPES = {
    # Process diagrams (linear flow)
    "process": {
        "layout_type": "process",
        "description": "Basic process flow (left to right)",
    },
    "chevron": {
        "layout_type": "chevron",
        "description": "Chevron/arrow process flow",
    },
    "arrow": {
        "layout_type": "arrow",
        "description": "Arrow process diagram",
    },
    # Cycle diagrams
    "cycle": {
        "layout_type": "cycle",
        "description": "Circular cycle diagram",
    },
    "radial": {
        "layout_type": "radial",
        "description": "Radial/hub-and-spoke diagram",
    },
    # Hierarchy diagrams
    "hierarchy": {
        "layout_type": "hierarchy",
        "description": "Organizational hierarchy",
    },
    "org_chart": {
        "layout_type": "orgChart",
        "description": "Organization chart",
    },
    # Relationship diagrams
    "venn": {
        "layout_type": "venn",
        "description": "Venn diagram (overlapping circles)",
    },
    "pyramid": {
        "layout_type": "pyramid",
        "description": "Pyramid diagram",
    },
    "matrix": {
        "layout_type": "matrix",
        "description": "2x2 or larger matrix",
    },
    # List diagrams
    "list": {
        "layout_type": "list",
        "description": "Vertical or horizontal list",
    },
    "target": {
        "layout_type": "target",
        "description": "Target/bullseye diagram",
    },
}


@dataclass
class DiagramNode:
    """A node in a diagram.

    Attributes:
        text: The node's text content
        children: Child nodes (for hierarchical diagrams)
        color: Optional color override (hex or theme color)
        description: Optional description text
    """

    text: str
    children: list[DiagramNode] = field(default_factory=list)
    color: str | None = None
    description: str | None = None

    @classmethod
    def from_dict(cls, data: dict[str, Any]) -> DiagramNode:
        """Create a DiagramNode from a dictionary."""
        children = []
        if "children" in data:
            children = [cls.from_dict(c) for c in data["children"]]
        return cls(
            text=data.get("text", ""),
            children=children,
            color=data.get("color"),
            description=data.get("description"),
        )


@dataclass
class DiagramData:
    """Data for a diagram.

    Attributes:
        nodes: List of top-level nodes
        title: Optional diagram title
    """

    nodes: list[DiagramNode] = field(default_factory=list)
    title: str | None = None

    @classmethod
    def from_list(cls, items: list[str | dict]) -> DiagramData:
        """Create DiagramData from a simple list.

        Accepts either strings or dicts with 'text' and optional 'children'.
        """
        nodes = []
        for item in items:
            if isinstance(item, str):
                nodes.append(DiagramNode(text=item))
            else:
                nodes.append(DiagramNode.from_dict(item))
        return cls(nodes=nodes)


@dataclass
class DiagramProperties:
    """Visual properties for a diagram.

    Attributes:
        diagram_type: Type of diagram (process, cycle, hierarchy, etc.)
        color_scheme: Color scheme name or theme reference
        style: Visual style (simple, 3d, etc.)
    """

    diagram_type: str = "process"
    color_scheme: str = "accent1"
    style: str = "simple"


class DiagramPart:
    """Handles diagram XML generation.

    A diagram in OOXML consists of multiple parts:
    - data.xml: The hierarchical data model
    - layout.xml: Layout definition
    - style.xml: Style information
    - colors.xml: Color mappings

    For simplicity, this generates a self-contained representation
    that can be embedded in a graphicFrame.
    """

    def __init__(
        self,
        data: DiagramData,
        props: DiagramProperties | None = None,
    ) -> None:
        self.data = data
        self.props = props or DiagramProperties()

    def to_data_xml(self) -> bytes:
        """Generate diagram data XML (dgm:dataModel)."""
        # Create root element
        root = etree.Element(
            qn("dgm:dataModel"),
            nsmap={
                "dgm": nsmap["dgm"],
                "a": nsmap["a"],
            },
        )

        # Point list
        pt_lst = etree.SubElement(root, qn("dgm:ptLst"))

        # Connection list
        cxn_lst = etree.SubElement(root, qn("dgm:cxnLst"))

        # Add document point (root of model)
        doc_pt = etree.SubElement(pt_lst, qn("dgm:pt"))
        doc_pt.set("modelId", "{0}")
        doc_pt.set("type", "doc")
        etree.SubElement(doc_pt, qn("dgm:prSet"))
        etree.SubElement(doc_pt, qn("dgm:spPr"))
        t = etree.SubElement(doc_pt, qn("dgm:t"))
        etree.SubElement(t, qn("a:bodyPr"))
        etree.SubElement(t, qn("a:lstStyle"))
        etree.SubElement(t, qn("a:p"))

        # Add data points
        self._add_nodes_to_model(
            pt_lst, cxn_lst, self.data.nodes, parent_id="{0}", start_id=1
        )

        return etree.tostring(root, xml_declaration=True, encoding="UTF-8")

    def _add_nodes_to_model(
        self,
        pt_lst: etree._Element,
        cxn_lst: etree._Element,
        nodes: list[DiagramNode],
        parent_id: str,
        start_id: int,
    ) -> int:
        """Add nodes to the data model recursively."""
        current_id = start_id

        for i, node in enumerate(nodes):
            node_id = f"{{{current_id}}}"
            current_id += 1

            # Create point
            pt = etree.SubElement(pt_lst, qn("dgm:pt"))
            pt.set("modelId", node_id)

            # Properties
            pr_set = etree.SubElement(pt, qn("dgm:prSet"))
            if node.color:
                # Set custom color
                pr_set.set("custLinFactNeighborX", "0")

            # Shape properties
            etree.SubElement(pt, qn("dgm:spPr"))

            # Text
            t = etree.SubElement(pt, qn("dgm:t"))
            etree.SubElement(t, qn("a:bodyPr"))
            etree.SubElement(t, qn("a:lstStyle"))
            p = etree.SubElement(t, qn("a:p"))
            r = etree.SubElement(p, qn("a:r"))
            a_t = etree.SubElement(r, qn("a:t"))
            a_t.text = node.text

            # Create connection from parent
            cxn = etree.SubElement(cxn_lst, qn("dgm:cxn"))
            cxn.set("modelId", f"{{{current_id}}}")
            current_id += 1
            cxn.set("srcId", parent_id)
            cxn.set("destId", node_id)
            cxn.set("srcOrd", str(i))
            cxn.set("destOrd", "0")
            cxn.set("parTransId", f"{{{current_id}}}")
            current_id += 1
            cxn.set("sibTransId", f"{{{current_id}}}")
            current_id += 1

            # Recursively add children
            if node.children:
                current_id = self._add_nodes_to_model(
                    pt_lst, cxn_lst, node.children, node_id, current_id
                )

        return current_id

    def to_layout_xml(self) -> bytes:
        """Generate diagram layout XML."""
        # For simplicity, use a basic layout definition
        root = etree.Element(
            qn("dgm:layoutDef"),
            nsmap={
                "dgm": nsmap["dgm"],
                "a": nsmap["a"],
            },
        )
        root.set("uniqueId", f"urn:microsoft.com/office/officeart/2005/8/layout/{self.props.diagram_type}")

        # Basic layout algorithm
        layout_node = etree.SubElement(root, qn("dgm:layoutNode"))
        layout_node.set("name", "root")

        # Algorithm
        alg = etree.SubElement(layout_node, qn("dgm:alg"))
        alg.set("type", "lin")  # linear

        # Shape
        shape = etree.SubElement(layout_node, qn("dgm:shape"))
        shape.set("type", "rect")

        # Presentation
        pres_of = etree.SubElement(layout_node, qn("dgm:presOf"))
        pres_of.set("axis", "self")

        return etree.tostring(root, xml_declaration=True, encoding="UTF-8")

    def to_style_xml(self) -> bytes:
        """Generate diagram style XML."""
        root = etree.Element(
            qn("dgm:styleDef"),
            nsmap={
                "dgm": nsmap["dgm"],
                "a": nsmap["a"],
            },
        )
        root.set("uniqueId", "urn:microsoft.com/office/officeart/2005/8/quickstyle/simple1")

        # Style label description
        title = etree.SubElement(root, qn("dgm:title"))
        title.set("val", "Simple")

        # Category list
        cat_lst = etree.SubElement(root, qn("dgm:catLst"))
        cat = etree.SubElement(cat_lst, qn("dgm:cat"))
        cat.set("type", "simple")
        cat.set("pri", "10000")

        return etree.tostring(root, xml_declaration=True, encoding="UTF-8")

    def to_colors_xml(self) -> bytes:
        """Generate diagram colors XML."""
        root = etree.Element(
            qn("dgm:colorsDef"),
            nsmap={
                "dgm": nsmap["dgm"],
                "a": nsmap["a"],
            },
        )
        root.set("uniqueId", f"urn:microsoft.com/office/officeart/2005/8/colors/{self.props.color_scheme}")

        # Title
        title = etree.SubElement(root, qn("dgm:title"))
        title.set("val", self.props.color_scheme.title())

        # Category list
        cat_lst = etree.SubElement(root, qn("dgm:catLst"))
        cat = etree.SubElement(cat_lst, qn("dgm:cat"))
        cat.set("type", "accent1")
        cat.set("pri", "10000")

        return etree.tostring(root, xml_declaration=True, encoding="UTF-8")


def create_diagram_graphic_frame(
    diagram_id: str,
    position: tuple[int, int, int, int],
    shape_id: int,
) -> etree._Element:
    """Create a graphicFrame element for a diagram.

    Args:
        diagram_id: Relationship ID for the diagram
        position: (x, y, cx, cy) in EMUs
        shape_id: Shape ID for this frame

    Returns:
        graphicFrame XML element
    """
    x, y, cx, cy = position

    # Create graphicFrame
    frame = etree.Element(qn("p:graphicFrame"))

    # Non-visual properties
    nv_graphic = etree.SubElement(frame, qn("p:nvGraphicFramePr"))
    c_nv_pr = etree.SubElement(nv_graphic, qn("p:cNvPr"))
    c_nv_pr.set("id", str(shape_id))
    c_nv_pr.set("name", f"Diagram {shape_id}")
    etree.SubElement(nv_graphic, qn("p:cNvGraphicFramePr"))
    nv_pr = etree.SubElement(nv_graphic, qn("p:nvPr"))  # noqa: F841

    # Transform
    xfrm = etree.SubElement(frame, qn("p:xfrm"))
    off = etree.SubElement(xfrm, qn("a:off"))
    off.set("x", str(x))
    off.set("y", str(y))
    ext = etree.SubElement(xfrm, qn("a:ext"))
    ext.set("cx", str(cx))
    ext.set("cy", str(cy))

    # Graphic
    graphic = etree.SubElement(frame, qn("a:graphic"))
    graphic_data = etree.SubElement(graphic, qn("a:graphicData"))
    graphic_data.set("uri", "http://schemas.openxmlformats.org/drawingml/2006/diagram")

    # Diagram relationship references
    rel_ids = etree.SubElement(
        graphic_data,
        qn("dgm:relIds"),
        nsmap={"dgm": nsmap["dgm"], "r": nsmap["r"]},
    )
    rel_ids.set(qn("r:dm"), f"{diagram_id}_data")
    rel_ids.set(qn("r:lo"), f"{diagram_id}_layout")
    rel_ids.set(qn("r:qs"), f"{diagram_id}_style")
    rel_ids.set(qn("r:cs"), f"{diagram_id}_colors")

    return frame


def get_diagram_types() -> dict[str, str]:
    """Get available diagram types with descriptions."""
    return {name: info["description"] for name, info in DIAGRAM_TYPES.items()}
