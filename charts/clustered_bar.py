"""
Clustered bar chart implementation for the pptx_charts_tables package.
"""

from pptx.chart.data import ChartData
from pptx.enum.chart import XL_CHART_TYPE
from pptx.dml.color import RGBColor
from .base import Chart


class ClusteredBarChart(Chart):
    """
    Clustered bar chart implementation.
    """

    def __init__(self, slide, data, position, size, config, **kwargs):
        """
        Initialize a clustered bar chart.

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
        """Create the clustered bar chart."""
        # Determine chart orientation (vertical or horizontal)
        chart_type = self.kwargs.get(
            "chart_type",
            self.config["clustered_bar_chart"].get("chart_type", "column_clustered"),
        )

        chart_types = {
            "column_clustered": XL_CHART_TYPE.COLUMN_CLUSTERED,  # Vertical bars
            "bar_clustered": XL_CHART_TYPE.BAR_CLUSTERED,  # Horizontal bars
        }

        xl_chart_type = chart_types.get(chart_type, XL_CHART_TYPE.COLUMN_CLUSTERED)

        # Prepare data
        chart_data = ChartData()

        # Get category column (first column by default)
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
        """Apply styles to the clustered bar chart."""
        # Set chart title
        self._set_chart_title()

        # Set gap width
        if hasattr(self.chart.plots[0], "gap_width"):
            gap_width = self.kwargs.get(
                "gap_width", self.config["clustered_bar_chart"].get("gap_width", 100)
            )
            self.chart.plots[0].gap_width = gap_width

        # Set up axes
        self._setup_category_axis()
        self._setup_value_axis()

        # Set up legend
        self._setup_legend()

        # Set up data labels
        self._setup_data_labels()

        # Set series colors
        self._set_series_colors()

    def _set_series_colors(self):
        """Set colors for chart series."""
        for i, series in enumerate(self.chart.series):
            series_name = series.name
            color_key = f"series_{i+1}"

            # Try to get color from kwargs first, then from config
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

            if color:
                fill = series.format.fill
                fill.solid()
                fill.fore_color.rgb = RGBColor.from_string(color)

                # Set data label colors for each series
                if self.kwargs.get(
                    "has_data_labels",
                    self.config["charts"].get("has_data_labels", True),
                ):
                    series.data_labels.font.color.rgb = RGBColor.from_string(color)

                    # Set data labels to show values
                    series.data_labels.show_value = True

                    # Hide zero values if configured
                    hide_zeros = self.kwargs.get("hide_zero_data_labels", True)
                    if hide_zeros:
                        for j, point in enumerate(series.points):
                            if series.values[j] == 0:
                                point.data_label.show_value = False
