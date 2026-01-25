"""Chart XML handling for PresentationML.

Charts in PowerPoint are stored as separate XML parts (ppt/charts/chartN.xml)
referenced from slides via relationships. This module handles creation and
manipulation of chart XML in DrawingML Chart format.

Supported chart types:
- Column (clustered, stacked, percentStacked)
- Bar (clustered, stacked, percentStacked)
- Line (with/without markers)
- Pie / Doughnut
- Area (standard, stacked, percentStacked)
- Scatter (markers, lines, or both)
"""

from __future__ import annotations

from dataclasses import dataclass, field
from typing import Literal

from lxml import etree

from .ns import NAMESPACES, qn

# Chart type mapping to DrawingML chart element names
CHART_TYPE_MAP = {
    # Column charts
    "column": "c:barChart",
    "clustered_column": "c:barChart",
    "stacked_column": "c:barChart",
    "percent_stacked_column": "c:barChart",
    # Bar charts (horizontal)
    "bar": "c:barChart",
    "clustered_bar": "c:barChart",
    "stacked_bar": "c:barChart",
    "percent_stacked_bar": "c:barChart",
    # Line charts
    "line": "c:lineChart",
    "line_markers": "c:lineChart",
    # Pie charts
    "pie": "c:pieChart",
    "doughnut": "c:doughnutChart",
    # Area charts
    "area": "c:areaChart",
    "stacked_area": "c:areaChart",
    "percent_stacked_area": "c:areaChart",
    # Scatter charts
    "scatter": "c:scatterChart",
    "scatter_lines": "c:scatterChart",
    "scatter_smooth": "c:scatterChart",
}

# Theme color scheme references for chart series
THEME_COLORS = [
    "accent1", "accent2", "accent3", "accent4", "accent5", "accent6",
]


@dataclass
class ChartSeries:
    """A data series in a chart.

    Attributes:
        name: Series name (shown in legend)
        values: List of numeric values
        color: Optional hex color (e.g., "#FF0000") or theme color (e.g., "accent1")
        categories: Optional category labels (for scatter charts with per-series categories)
    """

    name: str
    values: list[float | int]
    color: str | None = None
    categories: list[str | float] | None = None


@dataclass
class ChartData:
    """Data for creating a chart.

    Attributes:
        categories: Category labels (x-axis for most charts)
        series: List of data series
    """

    categories: list[str]
    series: list[ChartSeries] = field(default_factory=list)


@dataclass
class ChartProperties:
    """Chart display properties.

    Attributes:
        title: Chart title (None for no title)
        legend: Legend position ("right", "left", "top", "bottom", "none")
        data_labels: Show data labels on points
        stacked: For bar/column/area, stack series
        percent_stacked: Stack to 100%
        smooth: For line/scatter, use smooth lines
        markers: For line/scatter, show markers
        bar_direction: "col" for column, "bar" for horizontal bar
        hole_size: For doughnut, hole size percentage (0-90)
        vary_colors: For pie/doughnut, vary colors by point
    """

    title: str | None = None
    legend: Literal["right", "left", "top", "bottom", "none"] = "right"
    data_labels: bool = False
    stacked: bool = False
    percent_stacked: bool = False
    smooth: bool = False
    markers: bool = True
    bar_direction: Literal["col", "bar"] = "col"
    hole_size: int = 50
    vary_colors: bool = True


class ChartPart:
    """Handles chart XML (ppt/charts/chartN.xml).

    A chart part contains:
    - c:chartSpace: Root element
      - c:chart: Chart definition
        - c:plotArea: Plot area with chart type elements
        - c:legend: Legend (optional)
      - c:txPr: Text properties
    """

    def __init__(
        self,
        chart_type: str,
        data: ChartData,
        props: ChartProperties | None = None,
    ) -> None:
        """Initialize chart part.

        Args:
            chart_type: Chart type (e.g., "column", "bar", "line", "pie")
            data: Chart data with categories and series
            props: Chart display properties
        """
        self.chart_type = chart_type.lower().replace("-", "_").replace(" ", "_")
        self.data = data
        self.props = props or ChartProperties()
        self._element: etree._Element | None = None

    @classmethod
    def new(
        cls,
        chart_type: str,
        data: ChartData,
        props: ChartProperties | None = None,
    ) -> ChartPart:
        """Create a new chart part.

        Args:
            chart_type: Type of chart ("column", "bar", "line", "pie", "area", "scatter")
            data: Chart data
            props: Chart display properties

        Returns:
            ChartPart instance
        """
        return cls(chart_type, data, props)

    def to_xml(self) -> bytes:
        """Serialize chart to XML bytes.

        Returns:
            UTF-8 encoded XML bytes
        """
        root = self._build_chart_space()
        return etree.tostring(
            root,
            xml_declaration=True,
            encoding="UTF-8",
            standalone=True,
        )

    def _build_chart_space(self) -> etree._Element:
        """Build the c:chartSpace root element."""
        nsmap_chart = {
            None: NAMESPACES["c"],
            "a": NAMESPACES["a"],
            "r": NAMESPACES["r"],
        }

        chart_space = etree.Element(qn("c:chartSpace"), nsmap=nsmap_chart)

        # Add date1904 element (Excel compatibility)
        date1904 = etree.SubElement(chart_space, qn("c:date1904"))
        date1904.set("val", "0")

        # Add language
        lang = etree.SubElement(chart_space, qn("c:lang"))
        lang.set("val", "en-US")

        # Rounding mode
        round_mode = etree.SubElement(chart_space, qn("c:roundedCorners"))
        round_mode.set("val", "0")

        # Build chart element
        chart = etree.SubElement(chart_space, qn("c:chart"))
        self._build_chart(chart)

        # Add default text properties for theme consistency
        self._add_text_properties(chart_space)

        # External data reference (empty for embedded data)
        etree.SubElement(chart_space, qn("c:externalData"))

        return chart_space

    def _build_chart(self, chart: etree._Element) -> None:
        """Build the c:chart element with plot area and legend."""
        # Title
        if self.props.title:
            self._add_title(chart, self.props.title)
        else:
            auto_title = etree.SubElement(chart, qn("c:autoTitleDeleted"))
            auto_title.set("val", "1")

        # Plot area
        plot_area = etree.SubElement(chart, qn("c:plotArea"))

        # Layout (auto)
        etree.SubElement(plot_area, qn("c:layout"))

        # Add chart type element
        self._add_chart_type_element(plot_area)

        # Add axes (except for pie/doughnut)
        if self.chart_type not in ("pie", "doughnut"):
            self._add_axes(plot_area)

        # Legend
        if self.props.legend != "none":
            self._add_legend(chart)

        # Plot area visible border and fill
        plot_vis = etree.SubElement(chart, qn("c:plotVisOnly"))
        plot_vis.set("val", "1")

        disp_blanks = etree.SubElement(chart, qn("c:dispBlanksAs"))
        disp_blanks.set("val", "gap")

    def _add_title(self, chart: etree._Element, title_text: str) -> None:
        """Add chart title."""
        title = etree.SubElement(chart, qn("c:title"))

        tx = etree.SubElement(title, qn("c:tx"))
        rich = etree.SubElement(tx, qn("c:rich"))

        body_pr = etree.SubElement(rich, qn("a:bodyPr"))
        body_pr.set("rot", "0")
        body_pr.set("spcFirstLastPara", "1")
        body_pr.set("vertOverflow", "ellipsis")
        body_pr.set("vert", "horz")
        body_pr.set("wrap", "square")
        body_pr.set("anchor", "ctr")
        body_pr.set("anchorCtr", "1")

        etree.SubElement(rich, qn("a:lstStyle"))

        p = etree.SubElement(rich, qn("a:p"))
        ppr = etree.SubElement(p, qn("a:pPr"))
        def_rpr = etree.SubElement(ppr, qn("a:defRPr"))
        def_rpr.set("sz", "1400")
        def_rpr.set("b", "0")
        def_rpr.set("i", "0")

        r = etree.SubElement(p, qn("a:r"))
        rpr = etree.SubElement(r, qn("a:rPr"))
        rpr.set("lang", "en-US")
        t = etree.SubElement(r, qn("a:t"))
        t.text = title_text

        etree.SubElement(p, qn("a:endParaRPr"))

        etree.SubElement(title, qn("c:layout"))

        overlay = etree.SubElement(title, qn("c:overlay"))
        overlay.set("val", "0")

    def _add_chart_type_element(self, plot_area: etree._Element) -> None:
        """Add the appropriate chart type element."""
        chart_type = self.chart_type

        if chart_type in ("column", "clustered_column", "stacked_column",
                          "percent_stacked_column", "bar", "clustered_bar",
                          "stacked_bar", "percent_stacked_bar"):
            self._add_bar_chart(plot_area)
        elif chart_type in ("line", "line_markers"):
            self._add_line_chart(plot_area)
        elif chart_type == "pie":
            self._add_pie_chart(plot_area)
        elif chart_type == "doughnut":
            self._add_doughnut_chart(plot_area)
        elif chart_type in ("area", "stacked_area", "percent_stacked_area"):
            self._add_area_chart(plot_area)
        elif chart_type in ("scatter", "scatter_lines", "scatter_smooth"):
            self._add_scatter_chart(plot_area)
        else:
            # Default to column
            self._add_bar_chart(plot_area)

    def _add_bar_chart(self, plot_area: etree._Element) -> None:
        """Add bar/column chart element."""
        bar_chart = etree.SubElement(plot_area, qn("c:barChart"))

        # Bar direction (col = vertical/column, bar = horizontal/bar)
        bar_dir = etree.SubElement(bar_chart, qn("c:barDir"))
        if "bar" in self.chart_type and self.chart_type != "stacked_bar":
            if self.props.bar_direction == "bar" or "bar" in self.chart_type:
                bar_dir.set("val", "bar")
            else:
                bar_dir.set("val", "col")
        else:
            bar_dir.set("val", self.props.bar_direction)

        # Grouping
        grouping = etree.SubElement(bar_chart, qn("c:grouping"))
        if self.props.percent_stacked or "percent" in self.chart_type:
            grouping.set("val", "percentStacked")
        elif self.props.stacked or "stacked" in self.chart_type:
            grouping.set("val", "stacked")
        else:
            grouping.set("val", "clustered")

        # Vary colors (for single series)
        vary_colors = etree.SubElement(bar_chart, qn("c:varyColors"))
        vary_colors.set("val", "0")

        # Add series
        for idx, series in enumerate(self.data.series):
            self._add_bar_series(bar_chart, idx, series)

        # Data labels
        d_lbls = etree.SubElement(bar_chart, qn("c:dLbls"))
        show_val = etree.SubElement(d_lbls, qn("c:showVal"))
        show_val.set("val", "1" if self.props.data_labels else "0")
        show_cat = etree.SubElement(d_lbls, qn("c:showCatName"))
        show_cat.set("val", "0")
        show_ser = etree.SubElement(d_lbls, qn("c:showSerName"))
        show_ser.set("val", "0")
        show_pct = etree.SubElement(d_lbls, qn("c:showPercent"))
        show_pct.set("val", "0")

        # Gap width
        gap_width = etree.SubElement(bar_chart, qn("c:gapWidth"))
        gap_width.set("val", "150")

        # Overlap for stacked
        if self.props.stacked or self.props.percent_stacked or "stacked" in self.chart_type:
            overlap = etree.SubElement(bar_chart, qn("c:overlap"))
            overlap.set("val", "100")

        # Axis IDs
        ax_id_cat = etree.SubElement(bar_chart, qn("c:axId"))
        ax_id_cat.set("val", "100")
        ax_id_val = etree.SubElement(bar_chart, qn("c:axId"))
        ax_id_val.set("val", "200")

    def _add_bar_series(
        self,
        chart_elem: etree._Element,
        idx: int,
        series: ChartSeries,
    ) -> None:
        """Add a bar/column series."""
        ser = etree.SubElement(chart_elem, qn("c:ser"))

        # Index
        ser_idx = etree.SubElement(ser, qn("c:idx"))
        ser_idx.set("val", str(idx))
        order = etree.SubElement(ser, qn("c:order"))
        order.set("val", str(idx))

        # Series name
        tx = etree.SubElement(ser, qn("c:tx"))
        v = etree.SubElement(tx, qn("c:v"))
        v.text = series.name

        # Series color
        self._add_series_fill(ser, idx, series.color)

        # Categories
        cat = etree.SubElement(ser, qn("c:cat"))
        str_ref = etree.SubElement(cat, qn("c:strRef"))
        str_cache = etree.SubElement(str_ref, qn("c:strCache"))
        pt_count = etree.SubElement(str_cache, qn("c:ptCount"))
        pt_count.set("val", str(len(self.data.categories)))
        for i, cat_name in enumerate(self.data.categories):
            pt = etree.SubElement(str_cache, qn("c:pt"))
            pt.set("idx", str(i))
            v_elem = etree.SubElement(pt, qn("c:v"))
            v_elem.text = str(cat_name)

        # Values
        val = etree.SubElement(ser, qn("c:val"))
        num_ref = etree.SubElement(val, qn("c:numRef"))
        num_cache = etree.SubElement(num_ref, qn("c:numCache"))
        fmt_code = etree.SubElement(num_cache, qn("c:formatCode"))
        fmt_code.text = "General"
        pt_count = etree.SubElement(num_cache, qn("c:ptCount"))
        pt_count.set("val", str(len(series.values)))
        for i, value in enumerate(series.values):
            pt = etree.SubElement(num_cache, qn("c:pt"))
            pt.set("idx", str(i))
            v_elem = etree.SubElement(pt, qn("c:v"))
            v_elem.text = str(value) if value is not None else ""

    def _add_series_fill(
        self,
        ser: etree._Element,
        idx: int,
        color: str | None,
    ) -> None:
        """Add fill color to a series."""
        sp_pr = etree.SubElement(ser, qn("c:spPr"))
        solid_fill = etree.SubElement(sp_pr, qn("a:solidFill"))

        if color:
            if color.startswith("#"):
                # Hex color
                srgb = etree.SubElement(solid_fill, qn("a:srgbClr"))
                srgb.set("val", color.lstrip("#").upper())
            elif color.startswith("accent") or color in ("dk1", "dk2", "lt1", "lt2"):
                # Theme color
                scheme = etree.SubElement(solid_fill, qn("a:schemeClr"))
                scheme.set("val", color)
            else:
                # Assume hex without #
                srgb = etree.SubElement(solid_fill, qn("a:srgbClr"))
                srgb.set("val", color.upper())
        else:
            # Use theme accent color
            theme_color = THEME_COLORS[idx % len(THEME_COLORS)]
            scheme = etree.SubElement(solid_fill, qn("a:schemeClr"))
            scheme.set("val", theme_color)

    def _add_line_chart(self, plot_area: etree._Element) -> None:
        """Add line chart element."""
        line_chart = etree.SubElement(plot_area, qn("c:lineChart"))

        # Grouping
        grouping = etree.SubElement(line_chart, qn("c:grouping"))
        grouping.set("val", "standard")

        # Vary colors
        vary_colors = etree.SubElement(line_chart, qn("c:varyColors"))
        vary_colors.set("val", "0")

        # Add series
        for idx, series in enumerate(self.data.series):
            self._add_line_series(line_chart, idx, series)

        # Data labels
        d_lbls = etree.SubElement(line_chart, qn("c:dLbls"))
        show_val = etree.SubElement(d_lbls, qn("c:showVal"))
        show_val.set("val", "1" if self.props.data_labels else "0")
        show_cat = etree.SubElement(d_lbls, qn("c:showCatName"))
        show_cat.set("val", "0")
        show_ser = etree.SubElement(d_lbls, qn("c:showSerName"))
        show_ser.set("val", "0")
        show_pct = etree.SubElement(d_lbls, qn("c:showPercent"))
        show_pct.set("val", "0")

        # Markers
        marker = etree.SubElement(line_chart, qn("c:marker"))
        marker.set("val", "1" if self.props.markers else "0")

        # Smooth
        smooth = etree.SubElement(line_chart, qn("c:smooth"))
        smooth.set("val", "1" if self.props.smooth else "0")

        # Axis IDs
        ax_id_cat = etree.SubElement(line_chart, qn("c:axId"))
        ax_id_cat.set("val", "100")
        ax_id_val = etree.SubElement(line_chart, qn("c:axId"))
        ax_id_val.set("val", "200")

    def _add_line_series(
        self,
        chart_elem: etree._Element,
        idx: int,
        series: ChartSeries,
    ) -> None:
        """Add a line series."""
        ser = etree.SubElement(chart_elem, qn("c:ser"))

        # Index
        ser_idx = etree.SubElement(ser, qn("c:idx"))
        ser_idx.set("val", str(idx))
        order = etree.SubElement(ser, qn("c:order"))
        order.set("val", str(idx))

        # Series name
        tx = etree.SubElement(ser, qn("c:tx"))
        v = etree.SubElement(tx, qn("c:v"))
        v.text = series.name

        # Line color
        sp_pr = etree.SubElement(ser, qn("c:spPr"))
        ln = etree.SubElement(sp_pr, qn("a:ln"))
        ln.set("w", "28575")  # 2.25pt

        solid_fill = etree.SubElement(ln, qn("a:solidFill"))
        if series.color:
            if series.color.startswith("#"):
                srgb = etree.SubElement(solid_fill, qn("a:srgbClr"))
                srgb.set("val", series.color.lstrip("#").upper())
            else:
                scheme = etree.SubElement(solid_fill, qn("a:schemeClr"))
                scheme.set("val", series.color)
        else:
            theme_color = THEME_COLORS[idx % len(THEME_COLORS)]
            scheme = etree.SubElement(solid_fill, qn("a:schemeClr"))
            scheme.set("val", theme_color)

        # Marker
        marker = etree.SubElement(ser, qn("c:marker"))
        if self.props.markers:
            symbol = etree.SubElement(marker, qn("c:symbol"))
            symbol.set("val", "circle")
            size = etree.SubElement(marker, qn("c:size"))
            size.set("val", "5")
        else:
            symbol = etree.SubElement(marker, qn("c:symbol"))
            symbol.set("val", "none")

        # Categories
        cat = etree.SubElement(ser, qn("c:cat"))
        str_ref = etree.SubElement(cat, qn("c:strRef"))
        str_cache = etree.SubElement(str_ref, qn("c:strCache"))
        pt_count = etree.SubElement(str_cache, qn("c:ptCount"))
        pt_count.set("val", str(len(self.data.categories)))
        for i, cat_name in enumerate(self.data.categories):
            pt = etree.SubElement(str_cache, qn("c:pt"))
            pt.set("idx", str(i))
            v_elem = etree.SubElement(pt, qn("c:v"))
            v_elem.text = str(cat_name)

        # Values
        val = etree.SubElement(ser, qn("c:val"))
        num_ref = etree.SubElement(val, qn("c:numRef"))
        num_cache = etree.SubElement(num_ref, qn("c:numCache"))
        fmt_code = etree.SubElement(num_cache, qn("c:formatCode"))
        fmt_code.text = "General"
        pt_count = etree.SubElement(num_cache, qn("c:ptCount"))
        pt_count.set("val", str(len(series.values)))
        for i, value in enumerate(series.values):
            pt = etree.SubElement(num_cache, qn("c:pt"))
            pt.set("idx", str(i))
            v_elem = etree.SubElement(pt, qn("c:v"))
            v_elem.text = str(value) if value is not None else ""

        # Smooth
        smooth = etree.SubElement(ser, qn("c:smooth"))
        smooth.set("val", "1" if self.props.smooth else "0")

    def _add_pie_chart(self, plot_area: etree._Element) -> None:
        """Add pie chart element."""
        pie_chart = etree.SubElement(plot_area, qn("c:pieChart"))

        # Vary colors
        vary_colors = etree.SubElement(pie_chart, qn("c:varyColors"))
        vary_colors.set("val", "1" if self.props.vary_colors else "0")

        # Add series (typically just one for pie)
        if self.data.series:
            self._add_pie_series(pie_chart, 0, self.data.series[0])

        # Data labels
        d_lbls = etree.SubElement(pie_chart, qn("c:dLbls"))
        show_val = etree.SubElement(d_lbls, qn("c:showVal"))
        show_val.set("val", "1" if self.props.data_labels else "0")
        show_cat = etree.SubElement(d_lbls, qn("c:showCatName"))
        show_cat.set("val", "0")
        show_ser = etree.SubElement(d_lbls, qn("c:showSerName"))
        show_ser.set("val", "0")
        show_pct = etree.SubElement(d_lbls, qn("c:showPercent"))
        show_pct.set("val", "0")

        # First slice angle
        first_slice = etree.SubElement(pie_chart, qn("c:firstSliceAng"))
        first_slice.set("val", "0")

    def _add_pie_series(
        self,
        chart_elem: etree._Element,
        idx: int,
        series: ChartSeries,
    ) -> None:
        """Add a pie series."""
        ser = etree.SubElement(chart_elem, qn("c:ser"))

        # Index
        ser_idx = etree.SubElement(ser, qn("c:idx"))
        ser_idx.set("val", str(idx))
        order = etree.SubElement(ser, qn("c:order"))
        order.set("val", str(idx))

        # Series name
        tx = etree.SubElement(ser, qn("c:tx"))
        v = etree.SubElement(tx, qn("c:v"))
        v.text = series.name

        # Data point colors (vary by category)
        for i in range(len(self.data.categories)):
            d_pt = etree.SubElement(ser, qn("c:dPt"))
            pt_idx = etree.SubElement(d_pt, qn("c:idx"))
            pt_idx.set("val", str(i))
            sp_pr = etree.SubElement(d_pt, qn("c:spPr"))
            solid_fill = etree.SubElement(sp_pr, qn("a:solidFill"))
            theme_color = THEME_COLORS[i % len(THEME_COLORS)]
            scheme = etree.SubElement(solid_fill, qn("a:schemeClr"))
            scheme.set("val", theme_color)

        # Categories
        cat = etree.SubElement(ser, qn("c:cat"))
        str_ref = etree.SubElement(cat, qn("c:strRef"))
        str_cache = etree.SubElement(str_ref, qn("c:strCache"))
        pt_count = etree.SubElement(str_cache, qn("c:ptCount"))
        pt_count.set("val", str(len(self.data.categories)))
        for i, cat_name in enumerate(self.data.categories):
            pt = etree.SubElement(str_cache, qn("c:pt"))
            pt.set("idx", str(i))
            v_elem = etree.SubElement(pt, qn("c:v"))
            v_elem.text = str(cat_name)

        # Values
        val = etree.SubElement(ser, qn("c:val"))
        num_ref = etree.SubElement(val, qn("c:numRef"))
        num_cache = etree.SubElement(num_ref, qn("c:numCache"))
        fmt_code = etree.SubElement(num_cache, qn("c:formatCode"))
        fmt_code.text = "General"
        pt_count = etree.SubElement(num_cache, qn("c:ptCount"))
        pt_count.set("val", str(len(series.values)))
        for i, value in enumerate(series.values):
            pt = etree.SubElement(num_cache, qn("c:pt"))
            pt.set("idx", str(i))
            v_elem = etree.SubElement(pt, qn("c:v"))
            v_elem.text = str(value) if value is not None else ""

    def _add_doughnut_chart(self, plot_area: etree._Element) -> None:
        """Add doughnut chart element."""
        doughnut_chart = etree.SubElement(plot_area, qn("c:doughnutChart"))

        # Vary colors
        vary_colors = etree.SubElement(doughnut_chart, qn("c:varyColors"))
        vary_colors.set("val", "1" if self.props.vary_colors else "0")

        # Add series (typically just one for doughnut)
        if self.data.series:
            self._add_pie_series(doughnut_chart, 0, self.data.series[0])

        # Data labels
        d_lbls = etree.SubElement(doughnut_chart, qn("c:dLbls"))
        show_val = etree.SubElement(d_lbls, qn("c:showVal"))
        show_val.set("val", "1" if self.props.data_labels else "0")
        show_cat = etree.SubElement(d_lbls, qn("c:showCatName"))
        show_cat.set("val", "0")
        show_ser = etree.SubElement(d_lbls, qn("c:showSerName"))
        show_ser.set("val", "0")
        show_pct = etree.SubElement(d_lbls, qn("c:showPercent"))
        show_pct.set("val", "0")

        # First slice angle
        first_slice = etree.SubElement(doughnut_chart, qn("c:firstSliceAng"))
        first_slice.set("val", "0")

        # Hole size
        hole_size = etree.SubElement(doughnut_chart, qn("c:holeSize"))
        hole_size.set("val", str(self.props.hole_size))

    def _add_area_chart(self, plot_area: etree._Element) -> None:
        """Add area chart element."""
        area_chart = etree.SubElement(plot_area, qn("c:areaChart"))

        # Grouping
        grouping = etree.SubElement(area_chart, qn("c:grouping"))
        if self.props.percent_stacked or "percent" in self.chart_type:
            grouping.set("val", "percentStacked")
        elif self.props.stacked or "stacked" in self.chart_type:
            grouping.set("val", "stacked")
        else:
            grouping.set("val", "standard")

        # Vary colors
        vary_colors = etree.SubElement(area_chart, qn("c:varyColors"))
        vary_colors.set("val", "0")

        # Add series
        for idx, series in enumerate(self.data.series):
            self._add_area_series(area_chart, idx, series)

        # Data labels
        d_lbls = etree.SubElement(area_chart, qn("c:dLbls"))
        show_val = etree.SubElement(d_lbls, qn("c:showVal"))
        show_val.set("val", "1" if self.props.data_labels else "0")
        show_cat = etree.SubElement(d_lbls, qn("c:showCatName"))
        show_cat.set("val", "0")
        show_ser = etree.SubElement(d_lbls, qn("c:showSerName"))
        show_ser.set("val", "0")
        show_pct = etree.SubElement(d_lbls, qn("c:showPercent"))
        show_pct.set("val", "0")

        # Axis IDs
        ax_id_cat = etree.SubElement(area_chart, qn("c:axId"))
        ax_id_cat.set("val", "100")
        ax_id_val = etree.SubElement(area_chart, qn("c:axId"))
        ax_id_val.set("val", "200")

    def _add_area_series(
        self,
        chart_elem: etree._Element,
        idx: int,
        series: ChartSeries,
    ) -> None:
        """Add an area series."""
        ser = etree.SubElement(chart_elem, qn("c:ser"))

        # Index
        ser_idx = etree.SubElement(ser, qn("c:idx"))
        ser_idx.set("val", str(idx))
        order = etree.SubElement(ser, qn("c:order"))
        order.set("val", str(idx))

        # Series name
        tx = etree.SubElement(ser, qn("c:tx"))
        v = etree.SubElement(tx, qn("c:v"))
        v.text = series.name

        # Series fill color
        self._add_series_fill(ser, idx, series.color)

        # Categories
        cat = etree.SubElement(ser, qn("c:cat"))
        str_ref = etree.SubElement(cat, qn("c:strRef"))
        str_cache = etree.SubElement(str_ref, qn("c:strCache"))
        pt_count = etree.SubElement(str_cache, qn("c:ptCount"))
        pt_count.set("val", str(len(self.data.categories)))
        for i, cat_name in enumerate(self.data.categories):
            pt = etree.SubElement(str_cache, qn("c:pt"))
            pt.set("idx", str(i))
            v_elem = etree.SubElement(pt, qn("c:v"))
            v_elem.text = str(cat_name)

        # Values
        val = etree.SubElement(ser, qn("c:val"))
        num_ref = etree.SubElement(val, qn("c:numRef"))
        num_cache = etree.SubElement(num_ref, qn("c:numCache"))
        fmt_code = etree.SubElement(num_cache, qn("c:formatCode"))
        fmt_code.text = "General"
        pt_count = etree.SubElement(num_cache, qn("c:ptCount"))
        pt_count.set("val", str(len(series.values)))
        for i, value in enumerate(series.values):
            pt = etree.SubElement(num_cache, qn("c:pt"))
            pt.set("idx", str(i))
            v_elem = etree.SubElement(pt, qn("c:v"))
            v_elem.text = str(value) if value is not None else ""

    def _add_scatter_chart(self, plot_area: etree._Element) -> None:
        """Add scatter chart element."""
        scatter_chart = etree.SubElement(plot_area, qn("c:scatterChart"))

        # Scatter style
        scatter_style = etree.SubElement(scatter_chart, qn("c:scatterStyle"))
        if self.props.smooth or "smooth" in self.chart_type:
            scatter_style.set("val", "smoothMarker")
        elif "lines" in self.chart_type:
            scatter_style.set("val", "lineMarker")
        else:
            scatter_style.set("val", "marker")

        # Vary colors
        vary_colors = etree.SubElement(scatter_chart, qn("c:varyColors"))
        vary_colors.set("val", "0")

        # Add series
        for idx, series in enumerate(self.data.series):
            self._add_scatter_series(scatter_chart, idx, series)

        # Data labels
        d_lbls = etree.SubElement(scatter_chart, qn("c:dLbls"))
        show_val = etree.SubElement(d_lbls, qn("c:showVal"))
        show_val.set("val", "1" if self.props.data_labels else "0")
        show_cat = etree.SubElement(d_lbls, qn("c:showCatName"))
        show_cat.set("val", "0")
        show_ser = etree.SubElement(d_lbls, qn("c:showSerName"))
        show_ser.set("val", "0")
        show_pct = etree.SubElement(d_lbls, qn("c:showPercent"))
        show_pct.set("val", "0")

        # Axis IDs
        ax_id_x = etree.SubElement(scatter_chart, qn("c:axId"))
        ax_id_x.set("val", "100")
        ax_id_y = etree.SubElement(scatter_chart, qn("c:axId"))
        ax_id_y.set("val", "200")

    def _add_scatter_series(
        self,
        chart_elem: etree._Element,
        idx: int,
        series: ChartSeries,
    ) -> None:
        """Add a scatter series."""
        ser = etree.SubElement(chart_elem, qn("c:ser"))

        # Index
        ser_idx = etree.SubElement(ser, qn("c:idx"))
        ser_idx.set("val", str(idx))
        order = etree.SubElement(ser, qn("c:order"))
        order.set("val", str(idx))

        # Series name
        tx = etree.SubElement(ser, qn("c:tx"))
        v = etree.SubElement(tx, qn("c:v"))
        v.text = series.name

        # Series color (for line/marker)
        sp_pr = etree.SubElement(ser, qn("c:spPr"))
        ln = etree.SubElement(sp_pr, qn("a:ln"))
        ln.set("w", "28575")
        solid_fill = etree.SubElement(ln, qn("a:solidFill"))
        if series.color:
            if series.color.startswith("#"):
                srgb = etree.SubElement(solid_fill, qn("a:srgbClr"))
                srgb.set("val", series.color.lstrip("#").upper())
            else:
                scheme = etree.SubElement(solid_fill, qn("a:schemeClr"))
                scheme.set("val", series.color)
        else:
            theme_color = THEME_COLORS[idx % len(THEME_COLORS)]
            scheme = etree.SubElement(solid_fill, qn("a:schemeClr"))
            scheme.set("val", theme_color)

        # Marker
        marker = etree.SubElement(ser, qn("c:marker"))
        if self.props.markers:
            symbol = etree.SubElement(marker, qn("c:symbol"))
            symbol.set("val", "circle")
            size = etree.SubElement(marker, qn("c:size"))
            size.set("val", "5")
            # Marker fill
            sp_pr_marker = etree.SubElement(marker, qn("c:spPr"))
            solid_fill_marker = etree.SubElement(sp_pr_marker, qn("a:solidFill"))
            if series.color:
                if series.color.startswith("#"):
                    srgb = etree.SubElement(solid_fill_marker, qn("a:srgbClr"))
                    srgb.set("val", series.color.lstrip("#").upper())
                else:
                    scheme = etree.SubElement(solid_fill_marker, qn("a:schemeClr"))
                    scheme.set("val", series.color)
            else:
                theme_color = THEME_COLORS[idx % len(THEME_COLORS)]
                scheme = etree.SubElement(solid_fill_marker, qn("a:schemeClr"))
                scheme.set("val", theme_color)
        else:
            symbol = etree.SubElement(marker, qn("c:symbol"))
            symbol.set("val", "none")

        # X values (categories or numeric)
        x_val = etree.SubElement(ser, qn("c:xVal"))
        categories = series.categories if series.categories else self.data.categories

        # Check if categories are numeric
        all_numeric = all(isinstance(c, int | float) for c in categories)

        if all_numeric:
            num_ref = etree.SubElement(x_val, qn("c:numRef"))
            num_cache = etree.SubElement(num_ref, qn("c:numCache"))
            fmt_code = etree.SubElement(num_cache, qn("c:formatCode"))
            fmt_code.text = "General"
            pt_count = etree.SubElement(num_cache, qn("c:ptCount"))
            pt_count.set("val", str(len(categories)))
            for i, cat_val in enumerate(categories):
                pt = etree.SubElement(num_cache, qn("c:pt"))
                pt.set("idx", str(i))
                v_elem = etree.SubElement(pt, qn("c:v"))
                v_elem.text = str(cat_val)
        else:
            str_ref = etree.SubElement(x_val, qn("c:strRef"))
            str_cache = etree.SubElement(str_ref, qn("c:strCache"))
            pt_count = etree.SubElement(str_cache, qn("c:ptCount"))
            pt_count.set("val", str(len(categories)))
            for i, cat_name in enumerate(categories):
                pt = etree.SubElement(str_cache, qn("c:pt"))
                pt.set("idx", str(i))
                v_elem = etree.SubElement(pt, qn("c:v"))
                v_elem.text = str(cat_name)

        # Y values
        y_val = etree.SubElement(ser, qn("c:yVal"))
        num_ref = etree.SubElement(y_val, qn("c:numRef"))
        num_cache = etree.SubElement(num_ref, qn("c:numCache"))
        fmt_code = etree.SubElement(num_cache, qn("c:formatCode"))
        fmt_code.text = "General"
        pt_count = etree.SubElement(num_cache, qn("c:ptCount"))
        pt_count.set("val", str(len(series.values)))
        for i, value in enumerate(series.values):
            pt = etree.SubElement(num_cache, qn("c:pt"))
            pt.set("idx", str(i))
            v_elem = etree.SubElement(pt, qn("c:v"))
            v_elem.text = str(value) if value is not None else ""

        # Smooth
        smooth = etree.SubElement(ser, qn("c:smooth"))
        smooth.set("val", "1" if self.props.smooth or "smooth" in self.chart_type else "0")

    def _add_axes(self, plot_area: etree._Element) -> None:
        """Add category and value axes."""
        is_scatter = "scatter" in self.chart_type

        # Category axis (or X value axis for scatter)
        if is_scatter:
            cat_ax = etree.SubElement(plot_area, qn("c:valAx"))
        else:
            cat_ax = etree.SubElement(plot_area, qn("c:catAx"))

        ax_id = etree.SubElement(cat_ax, qn("c:axId"))
        ax_id.set("val", "100")

        scaling = etree.SubElement(cat_ax, qn("c:scaling"))
        orientation = etree.SubElement(scaling, qn("c:orientation"))
        orientation.set("val", "minMax")

        delete = etree.SubElement(cat_ax, qn("c:delete"))
        delete.set("val", "0")

        ax_pos = etree.SubElement(cat_ax, qn("c:axPos"))
        ax_pos.set("val", "b")  # bottom

        # Major gridlines
        major_gridlines = etree.SubElement(cat_ax, qn("c:majorGridlines"))
        sp_pr = etree.SubElement(major_gridlines, qn("c:spPr"))
        ln = etree.SubElement(sp_pr, qn("a:ln"))
        ln.set("w", "9525")
        solid_fill = etree.SubElement(ln, qn("a:solidFill"))
        scheme_clr = etree.SubElement(solid_fill, qn("a:schemeClr"))
        scheme_clr.set("val", "tx1")
        lum_mod = etree.SubElement(scheme_clr, qn("a:lumMod"))
        lum_mod.set("val", "15000")
        lum_off = etree.SubElement(scheme_clr, qn("a:lumOff"))
        lum_off.set("val", "85000")

        if not is_scatter:
            num_fmt = etree.SubElement(cat_ax, qn("c:numFmt"))
            num_fmt.set("formatCode", "General")
            num_fmt.set("sourceLinked", "1")

        major_tick = etree.SubElement(cat_ax, qn("c:majorTickMark"))
        major_tick.set("val", "none")

        minor_tick = etree.SubElement(cat_ax, qn("c:minorTickMark"))
        minor_tick.set("val", "none")

        tick_lbl_pos = etree.SubElement(cat_ax, qn("c:tickLblPos"))
        tick_lbl_pos.set("val", "nextTo")

        cross_ax = etree.SubElement(cat_ax, qn("c:crossAx"))
        cross_ax.set("val", "200")

        crosses = etree.SubElement(cat_ax, qn("c:crosses"))
        crosses.set("val", "autoZero")

        if not is_scatter:
            auto = etree.SubElement(cat_ax, qn("c:auto"))
            auto.set("val", "1")

            lbl_algn = etree.SubElement(cat_ax, qn("c:lblAlgn"))
            lbl_algn.set("val", "ctr")

            lbl_offset = etree.SubElement(cat_ax, qn("c:lblOffset"))
            lbl_offset.set("val", "100")

        # Value axis
        val_ax = etree.SubElement(plot_area, qn("c:valAx"))

        ax_id = etree.SubElement(val_ax, qn("c:axId"))
        ax_id.set("val", "200")

        scaling = etree.SubElement(val_ax, qn("c:scaling"))
        orientation = etree.SubElement(scaling, qn("c:orientation"))
        orientation.set("val", "minMax")

        delete = etree.SubElement(val_ax, qn("c:delete"))
        delete.set("val", "0")

        ax_pos = etree.SubElement(val_ax, qn("c:axPos"))
        ax_pos.set("val", "l")  # left

        major_gridlines = etree.SubElement(val_ax, qn("c:majorGridlines"))
        sp_pr = etree.SubElement(major_gridlines, qn("c:spPr"))
        ln = etree.SubElement(sp_pr, qn("a:ln"))
        ln.set("w", "9525")
        solid_fill = etree.SubElement(ln, qn("a:solidFill"))
        scheme_clr = etree.SubElement(solid_fill, qn("a:schemeClr"))
        scheme_clr.set("val", "tx1")
        lum_mod = etree.SubElement(scheme_clr, qn("a:lumMod"))
        lum_mod.set("val", "15000")
        lum_off = etree.SubElement(scheme_clr, qn("a:lumOff"))
        lum_off.set("val", "85000")

        num_fmt = etree.SubElement(val_ax, qn("c:numFmt"))
        num_fmt.set("formatCode", "General")
        num_fmt.set("sourceLinked", "1")

        major_tick = etree.SubElement(val_ax, qn("c:majorTickMark"))
        major_tick.set("val", "none")

        minor_tick = etree.SubElement(val_ax, qn("c:minorTickMark"))
        minor_tick.set("val", "none")

        tick_lbl_pos = etree.SubElement(val_ax, qn("c:tickLblPos"))
        tick_lbl_pos.set("val", "nextTo")

        cross_ax = etree.SubElement(val_ax, qn("c:crossAx"))
        cross_ax.set("val", "100")

        crosses = etree.SubElement(val_ax, qn("c:crosses"))
        crosses.set("val", "autoZero")

        cross_between = etree.SubElement(val_ax, qn("c:crossBetween"))
        cross_between.set("val", "between")

    def _add_legend(self, chart: etree._Element) -> None:
        """Add chart legend."""
        legend = etree.SubElement(chart, qn("c:legend"))

        legend_pos = etree.SubElement(legend, qn("c:legendPos"))
        pos_map = {"right": "r", "left": "l", "top": "t", "bottom": "b"}
        legend_pos.set("val", pos_map.get(self.props.legend, "r"))

        etree.SubElement(legend, qn("c:layout"))

        overlay = etree.SubElement(legend, qn("c:overlay"))
        overlay.set("val", "0")

    def _add_text_properties(self, chart_space: etree._Element) -> None:
        """Add default text properties for theme consistency."""
        tx_pr = etree.SubElement(chart_space, qn("c:txPr"))

        body_pr = etree.SubElement(tx_pr, qn("a:bodyPr"))
        body_pr.set("rot", "0")
        body_pr.set("spcFirstLastPara", "1")
        body_pr.set("vertOverflow", "ellipsis")
        body_pr.set("vert", "horz")
        body_pr.set("wrap", "square")
        body_pr.set("anchor", "ctr")
        body_pr.set("anchorCtr", "1")

        etree.SubElement(tx_pr, qn("a:lstStyle"))

        p = etree.SubElement(tx_pr, qn("a:p"))
        ppr = etree.SubElement(p, qn("a:pPr"))
        def_rpr = etree.SubElement(ppr, qn("a:defRPr"))
        def_rpr.set("sz", "1000")
        def_rpr.set("b", "0")
        def_rpr.set("i", "0")
        def_rpr.set("u", "none")
        def_rpr.set("strike", "noStrike")
        def_rpr.set("kern", "1200")
        def_rpr.set("baseline", "0")

        # Use theme font
        latin = etree.SubElement(def_rpr, qn("a:latin"))
        latin.set("typeface", "+mn-lt")  # Minor latin (body font)
        ea = etree.SubElement(def_rpr, qn("a:ea"))
        ea.set("typeface", "+mn-ea")
        cs = etree.SubElement(def_rpr, qn("a:cs"))
        cs.set("typeface", "+mn-cs")

        etree.SubElement(p, qn("a:endParaRPr"))

    @classmethod
    def from_xml(cls, xml_bytes: bytes) -> ChartPart:
        """Parse chart from XML bytes.

        Args:
            xml_bytes: UTF-8 encoded XML

        Returns:
            ChartPart instance
        """
        root = etree.fromstring(xml_bytes)
        chart = root.find(qn("c:chart"))
        if chart is None:
            raise ValueError("Invalid chart XML: missing c:chart element")

        # Determine chart type
        plot_area = chart.find(qn("c:plotArea"))
        chart_type = "column"  # default

        if plot_area is not None:
            if plot_area.find(qn("c:barChart")) is not None:
                bar_chart = plot_area.find(qn("c:barChart"))
                bar_dir = bar_chart.find(qn("c:barDir"))
                if bar_dir is not None and bar_dir.get("val") == "bar":
                    chart_type = "bar"
                else:
                    chart_type = "column"
            elif plot_area.find(qn("c:lineChart")) is not None:
                chart_type = "line"
            elif plot_area.find(qn("c:pieChart")) is not None:
                chart_type = "pie"
            elif plot_area.find(qn("c:doughnutChart")) is not None:
                chart_type = "doughnut"
            elif plot_area.find(qn("c:areaChart")) is not None:
                chart_type = "area"
            elif plot_area.find(qn("c:scatterChart")) is not None:
                chart_type = "scatter"

        # Extract data (simplified - real implementation would parse fully)
        data = ChartData(categories=[], series=[])

        # Create instance
        instance = cls(chart_type, data, ChartProperties())
        instance._element = root
        return instance
