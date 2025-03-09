"""
Line chart implementation for the pptx_charts_tables package.
"""

from pptx.chart.data import ChartData
from pptx.enum.chart import XL_CHART_TYPE
from pptx.enum.dml import MSO_LINE
from pptx.dml.color import RGBColor
from .base import Chart


class LineChart(Chart):
    """
    Line chart implementation.
    """

    def __init__(self, slide, data, position, size, config, **kwargs):
        """
        Initialize a line chart.

        Args:
            slide (PPTXSlide): Parent slide.
            data (pd.DataFrame): Data for the chart.
            position (tuple): (x, y) position in inches.
            size (tuple): (width, height) size in inches.
            config (dict): Configuration.
            **kwargs: Additional chart-specific options.
        """
        super().__init__(slide, data, position, size, config, **kwargs)
        self._create_chart()
        self._apply_styles()

    def _create_chart(self):
        """Create the line chart."""
        # Determine chart type
        chart_type = self.kwargs.get(
            "chart_type", self.config["line_chart"].get("chart_type", "line")
        )

        # Check if we want markers
        show_markers = self.kwargs.get(
            "show_markers", self.config["line_chart"].get("show_markers", False)
        )

        # Choose the right chart type
        chart_types = {
            "line": (
                XL_CHART_TYPE.LINE if not show_markers else XL_CHART_TYPE.LINE_MARKERS
            ),
            "line_markers": XL_CHART_TYPE.LINE_MARKERS,
            "line_stacked": (
                XL_CHART_TYPE.LINE_STACKED
                if not show_markers
                else XL_CHART_TYPE.LINE_MARKERS_STACKED
            ),
            "line_stacked_markers": XL_CHART_TYPE.LINE_MARKERS_STACKED,
            "line_stacked_100": (
                XL_CHART_TYPE.LINE_STACKED_100
                if not show_markers
                else XL_CHART_TYPE.LINE_MARKERS_STACKED_100
            ),
            "line_stacked_100_markers": XL_CHART_TYPE.LINE_MARKERS_STACKED_100,
        }

        xl_chart_type = chart_types.get(chart_type, XL_CHART_TYPE.LINE)

        # Prepare data
        chart_data = ChartData()

        # Get category column (first column by default, often dates/time periods)
        category_col = self.kwargs.get("category_column", self.data.columns[0])

        # Get series columns (all except category column by default)
        series_cols = self.kwargs.get(
            "series_columns", [col for col in self.data.columns if col != category_col]
        )

        # Add categories
        chart_data.categories = self.data[category_col]

        # Add series
        for i, col in enumerate(series_cols):
            # Use series name from kwargs if provided, otherwise use column name
            series_name = self.kwargs.get(f"series_{i+1}_name", col)
            chart_data.add_series(series_name, self.data[col])

        # Create chart
        self.chart = self.slide.slide.shapes.add_chart(
            xl_chart_type,
            self.position[0],
            self.position[1],
            self.size[0],
            self.size[1],
            chart_data,
        ).chart

    def _apply_styles(self):
        """Apply styles to the line chart."""
        # Set chart title
        self._set_chart_title()

        # Line charts typically show axes
        self.kwargs["value_axis_visible"] = self.kwargs.get(
            "value_axis_visible",
            self.config["line_chart"].get("value_axis_visible", True),
        )

        # Line charts often show gridlines
        self.kwargs["value_axis_has_gridlines"] = self.kwargs.get(
            "value_axis_has_gridlines",
            self.config["line_chart"].get("value_axis_has_gridlines", True),
        )

        # Set up axes
        self._setup_category_axis()
        self._setup_value_axis()

        # Configure gridlines
        self._configure_gridlines()

        # Set up legend
        self._setup_legend()

        # Set up data labels (not as common for line charts)
        if self.kwargs.get("has_data_labels", False):
            self._setup_data_labels()

        # Style the lines
        self._style_lines()

    def _configure_gridlines(self):
        """Configure gridlines for line chart."""
        if not hasattr(self.chart, "value_axis"):
            return

        if not self.chart.value_axis.has_major_gridlines:
            return

        # Set gridline style
        gridline_color = self.kwargs.get(
            "gridline_color", self.config["line_chart"].get("gridline_color", "BFBFBF")
        )

        gridline_dash_style = self.kwargs.get(
            "gridline_dash_style",
            self.config["line_chart"].get("gridline_dash_style", "dash"),
        )

        dash_styles = {
            "dash": MSO_LINE.DASH,
            "round_dot": MSO_LINE.ROUND_DOT,
            "square_dot": MSO_LINE.SQUARE_DOT,
            "dash_dot": MSO_LINE.DASH_DOT,
            "long_dash": MSO_LINE.LONG_DASH,
            "long_dash_dot": MSO_LINE.LONG_DASH_DOT,
            "solid": MSO_LINE.SOLID,
        }

        gridlines = self.chart.value_axis.major_gridlines

        if gridline_color:
            gridlines.format.line.color.rgb = RGBColor.from_string(gridline_color)

        if gridline_dash_style in dash_styles:
            gridlines.format.line.dash_style = dash_styles[gridline_dash_style]

        # Set gridline width
        gridline_width = self.kwargs.get("gridline_width", 0.5)
        if hasattr(gridlines.format.line, "width"):
            # python-pptx uses EMU units (English Metric Units)
            # 1 point = 12700 EMU
            gridlines.format.line.width = int(gridline_width * 12700)

    def _style_lines(self):
        """Style the lines in the line chart."""
        # Line styles for each series
        line_styles = {
            "solid": MSO_LINE.SOLID,
            "dash": MSO_LINE.DASH,
            "round_dot": MSO_LINE.ROUND_DOT,
            "square_dot": MSO_LINE.SQUARE_DOT,
            "dash_dot": MSO_LINE.DASH_DOT,
            "long_dash": MSO_LINE.LONG_DASH,
            "long_dash_dot": MSO_LINE.LONG_DASH_DOT,
        }

        for i, series in enumerate(self.chart.series):
            series_name = series.name
            color_key = f"series_{i+1}"

            # Set line color
            color = self.kwargs.get(
                f"{series_name}_color",
                self.kwargs.get(
                    f"series_{i+1}_color",
                    self.config["charts"]["colors"].get(
                        series_name,
                        self.config["charts"]["colors"].get(color_key, None),
                    ),
                ),
            )

            # Set line width
            width = self.kwargs.get(
                f"{series_name}_line_width",
                self.kwargs.get(
                    f"series_{i+1}_line_width",
                    self.config["line_chart"]["line_width"].get(
                        series_name,
                        self.config["line_chart"]["line_width"].get(
                            color_key, 1.5  # Default line width
                        ),
                    ),
                ),
            )

            # Set line style
            style = self.kwargs.get(
                f"{series_name}_line_style",
                self.kwargs.get(
                    f"series_{i+1}_line_style",
                    self.config["line_chart"]["line_style"].get(
                        series_name,
                        self.config["line_chart"]["line_style"].get(
                            color_key, "solid"  # Default line style
                        ),
                    ),
                ),
            )

            line = series.format.line

            if color:
                line.color.rgb = RGBColor.from_string(color)

            if style in line_styles:
                line.dash_style = line_styles[style]

            # python-pptx uses EMU units (English Metric Units)
            # 1 point = 12700 EMU
            if width:
                line.width = int(width * 12700)
