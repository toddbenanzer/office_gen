"""
Donut chart implementation for the pptx_charts_tables package.
"""

from pptx.chart.data import CategoryChartData
from pptx.enum.chart import XL_CHART_TYPE
from pptx.dml.color import RGBColor
from .base import Chart


class DonutChart(Chart):
    """
    Donut chart implementation.
    """

    def __init__(self, slide, data, position, size, config, **kwargs):
        """
        Initialize a donut chart.

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
        """Create the donut chart."""
        # Determine chart type (standard or exploded)
        chart_type = self.kwargs.get(
            "chart_type", self.config["donut_chart"].get("chart_type", "doughnut")
        )

        chart_types = {
            "doughnut": XL_CHART_TYPE.DOUGHNUT,
            "doughnut_exploded": XL_CHART_TYPE.DOUGHNUT_EXPLODED,
        }

        xl_chart_type = chart_types.get(chart_type, XL_CHART_TYPE.DOUGHNUT)

        # Prepare data
        chart_data = CategoryChartData()

        # For donut chart, data structure is slightly different:
        # - The category column becomes the labels
        # - We only use one series column for the values

        # Get category column (first column by default)
        category_col = self.kwargs.get("category_column", self.data.columns[0])

        # Get value column (second column by default or specified)
        value_col = self.kwargs.get(
            "value_column", self.data.columns[1] if len(self.data.columns) > 1 else None
        )

        if value_col is None:
            raise ValueError("No value column available for donut chart")

        # Add categories
        chart_data.categories = self.data[category_col]

        # Add the single series
        series_name = self.kwargs.get("series_name", value_col)
        chart_data.add_series(series_name, self.data[value_col])

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
        """Apply styles to the donut chart."""
        # Set chart title
        self._set_chart_title()

        # Set up legend
        self._setup_legend()

        # Set up data labels with percentage format by default
        self._setup_data_labels()
        self._set_data_label_format()

        # Set segment colors
        self._set_segment_colors()

    def _set_data_label_format(self):
        """Set data label format for donut chart."""
        if not hasattr(self.chart, "plots") or not self.chart.plots:
            return

        plot = self.chart.plots[0]

        if not plot.has_data_labels:
            return

        data_labels = plot.data_labels

        # Set number format (percentage by default for donut charts)
        number_format = self.kwargs.get(
            "data_label_number_format",
            self.config["donut_chart"].get("data_label_number_format", "0%"),
        )

        data_labels.number_format = number_format

    def _set_segment_colors(self):
        """Set colors for donut chart segments."""
        # For donut charts, we need to set colors for individual points
        series = self.chart.series[0]
        category_col = self.kwargs.get("category_column", self.data.columns[0])

        # Get explicit colors for each segment if provided
        segment_colors = self.kwargs.get("segment_colors", {})

        # Get the category column values
        categories = self.data[category_col].tolist()

        # Set colors for each segment
        for i, point in enumerate(series.points):
            category = categories[i] if i < len(categories) else f"Category {i+1}"

            # Try to get color from segment_colors, then from normal config
            color = segment_colors.get(
                category, self.config["charts"]["colors"].get(f"series_{i+1}", None)
            )

            if color:
                point.format.fill.solid()
                point.format.fill.fore_color.rgb = RGBColor.from_string(color)
