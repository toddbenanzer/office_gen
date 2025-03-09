"""
Example usage of the pptx_charts_tables package.
"""

import pandas as pd
from pptx_charts_tables import PPTXPresentation


def main():
    # Create sample data
    quarterly_data = pd.DataFrame(
        {
            "quarter": ["Q1", "Q2", "Q3", "Q4"],
            "revenue": [1250000, 1450000, 1550000, 1750000],
            "costs": [950000, 1050000, 1150000, 1250000],
            "profit": [300000, 400000, 400000, 500000],
        }
    )

    product_data = pd.DataFrame(
        {
            "product": ["Product A", "Product B", "Product C", "Product D"],
            "sales": [352, 284, 156, 108],
        }
    )

    trend_data = pd.DataFrame(
        {
            "month": [
                "Jan",
                "Feb",
                "Mar",
                "Apr",
                "May",
                "Jun",
                "Jul",
                "Aug",
                "Sep",
                "Oct",
                "Nov",
                "Dec",
            ],
            "this_year": [45, 50, 55, 59, 65, 70, 75, 80, 85, 89, 91, 95],
            "last_year": [40, 43, 45, 50, 55, 59, 65, 68, 72, 75, 78, 80],
        }
    )

    segment_data = pd.DataFrame(
        {
            "segment": ["Enterprise", "Mid-Market", "SMB", "Direct"],
            "revenue": [2500000, 1500000, 750000, 250000],
        }
    )

    # Create the presentation
    prs = PPTXPresentation()

    # Example 1: Bar Chart Slide
    slide1 = prs.add_slide(title="Quarterly Financial Performance")

    # Add a clustered bar chart
    slide1.add_chart(
        chart_type="clustered_bar",
        data=quarterly_data,
        position=(1, 1.5),
        size=(8, 4.5),
        title="Quarterly Revenue, Costs, and Profit",
        category_column="quarter",
        series_columns=["revenue", "costs", "profit"],
        series_1_name="Revenue",
        series_2_name="Costs",
        series_3_name="Profit",
        series_1_color="4472C4",  # Blue
        series_2_color="ED7D31",  # Orange
        series_3_color="70AD47",  # Green
        value_axis_visible=True,
        has_data_labels=True,
        data_label_number_format="$#,##0,K",
    )

    # Add a table below the chart
    slide1.add_table(
        data=quarterly_data,
        position=(1, 6.25),
        table_width=8,
        column_formats={
            "quarter": "text",
            "revenue": "dollars",
            "costs": "dollars",
            "profit": "dollars",
        },
        col_widths=[1.5, 2.17, 2.17, 2.16],  # Specify exact column widths
        has_header=True,
        header_style={"font_color": "FFFFFF", "fill_color": "4472C4"},
        total_rows=[4],  # Add a total row at the end
    )

    # Example 2: Donut Chart Slide
    slide2 = prs.add_slide(title="Product Sales Distribution")

    # Add a donut chart
    slide2.add_chart(
        chart_type="donut",
        data=product_data,
        position=(1, 1.5),
        size=(4, 4.5),
        title="Product Sales Breakdown",
        category_column="product",
        value_column="sales",
        series_name="Sales",
        has_legend=True,
        legend_position="right",
        segment_colors={
            "Product A": "4472C4",  # Blue
            "Product B": "5B9BD5",  # Light Blue
            "Product C": "8FAADC",  # Pale Blue
            "Product D": "B4C7E7",  # Very Pale Blue
        },
        data_label_number_format="0%",
    )

    # Add a table with the data
    slide2.add_table(
        data=product_data,
        position=(5.5, 3),
        table_width=4,
        column_formats={"product": "text", "sales": "counts"},
        alternating_row_fill=True,
        alternating_row_fill_color="E6E6E6",
    )

    # Example 3: Line Chart Slide
    slide3 = prs.add_slide(title="Monthly Performance Trends")

    # Add a line chart
    slide3.add_chart(
        chart_type="line",
        data=trend_data,
        position=(1, 1.5),
        size=(8, 4.5),
        title="Monthly Performance Comparison",
        category_column="month",
        series_columns=["this_year", "last_year"],
        series_1_name="This Year",
        series_2_name="Last Year",
        series_1_color="4472C4",  # Blue
        series_2_color="ED7D31",  # Orange
        series_1_line_width=2.5,
        series_2_line_width=1.5,
        series_1_line_style="solid",
        series_2_line_style="dash",
        show_markers=True,
        value_axis_visible=True,
        value_axis_has_gridlines=True,
        gridline_color="E6E6E6",
        gridline_dash_style="dash",
    )

    # Example 4: Stacked Bar Chart Slide
    slide4 = prs.add_slide(title="Revenue by Segment")

    # Add a stacked bar chart
    slide4.add_chart(
        chart_type="stacked_bar",
        data=segment_data,
        position=(1, 1.5),
        size=(8, 4.5),
        title="Revenue Distribution by Segment",
        category_column="segment",
        series_columns=["revenue"],
        series_1_name="Revenue",
        series_1_color="4472C4",  # Blue
        chart_type="column_stacked",
        has_data_labels=True,
        data_label_position="inside_end",
        data_label_font_color="FFFFFF",
        data_label_number_format="$#,##0,K",
    )

    # Save the presentation
    prs.save("financial_report.pptx")
    print("Presentation created: financial_report.pptx")


if __name__ == "__main__":
    main()
