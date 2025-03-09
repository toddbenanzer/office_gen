"""
Advanced financial example for the pptx_charts_tables package.
This example demonstrates creating a comprehensive financial report presentation.
"""

import pandas as pd
import numpy as np
from pptx_charts_tables import PPTXPresentation


def create_financial_report():
    """Create a comprehensive financial report presentation."""

    # Set up sample data
    # Monthly financial data for past year
    months = [
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
    ]

    # Revenue data with seasonal pattern
    base_revenue = 1000000  # $1M base monthly revenue
    revenue_data = [
        base_revenue * (1 + 0.05 * i + 0.1 * np.sin(np.pi * i / 6)) for i in range(12)
    ]

    # Cost data (70-80% of revenue with random variation)
    cost_data = [rev * (0.7 + 0.1 * np.random.random()) for rev in revenue_data]

    # Marketing expenses (10-15% of revenue)
    marketing_data = [rev * (0.1 + 0.05 * np.random.random()) for rev in revenue_data]

    # R&D expenses (5-10% of revenue)
    rd_data = [rev * (0.05 + 0.05 * np.random.random()) for rev in revenue_data]

    # Administrative expenses (8-12% of revenue)
    admin_data = [rev * (0.08 + 0.04 * np.random.random()) for rev in revenue_data]

    # Calculate profits
    profit_data = [
        revenue_data[i] - cost_data[i] - marketing_data[i] - rd_data[i] - admin_data[i]
        for i in range(12)
    ]

    # Calculate profit margins
    margin_data = [
        profit / revenue * 100 if revenue > 0 else 0
        for profit, revenue in zip(profit_data, revenue_data)
    ]

    # Create pandas DataFrames
    monthly_data = pd.DataFrame(
        {
            "Month": months,
            "Revenue": revenue_data,
            "Costs": cost_data,
            "Marketing": marketing_data,
            "R&D": rd_data,
            "Admin": admin_data,
            "Profit": profit_data,
            "Margin": margin_data,
        }
    )

    # Create quarterly data
    quarterly_data = pd.DataFrame(
        {
            "Quarter": ["Q1", "Q2", "Q3", "Q4"],
            "Revenue": [
                sum(revenue_data[0:3]),
                sum(revenue_data[3:6]),
                sum(revenue_data[6:9]),
                sum(revenue_data[9:12]),
            ],
            "Costs": [
                sum(cost_data[0:3]),
                sum(cost_data[3:6]),
                sum(cost_data[6:9]),
                sum(cost_data[9:12]),
            ],
            "Profit": [
                sum(profit_data[0:3]),
                sum(profit_data[3:6]),
                sum(profit_data[6:9]),
                sum(profit_data[9:12]),
            ],
        }
    )

    # Add YoY Growth column (random values between 5% and 20%)
    quarterly_data["YoY_Growth"] = [5 + 15 * np.random.random() for _ in range(4)]

    # Create expense breakdown data
    expense_categories = ["COGS", "Marketing", "R&D", "Administration", "Other"]
    expense_values = [
        sum(cost_data),
        sum(marketing_data),
        sum(rd_data),
        sum(admin_data),
        sum(revenue_data) * 0.02,  # Other expenses (2% of revenue)
    ]

    expense_data = pd.DataFrame(
        {"Category": expense_categories, "Amount": expense_values}
    )

    # Product performance data
    products = ["Product A", "Product B", "Product C", "Product D", "Product E"]
    product_revenue = [4500000, 3200000, 2100000, 1800000, 900000]
    product_growth = [18.5, 12.3, -5.2, 22.1, 8.7]
    product_margin = [42.3, 38.7, 29.4, 35.2, 44.1]

    product_data = pd.DataFrame(
        {
            "Product": products,
            "Revenue": product_revenue,
            "Growth": product_growth,
            "Margin": product_margin,
        }
    )

    # Geographic data
    regions = ["North America", "Europe", "Asia Pacific", "Latin America", "MEA"]
    region_revenue = [5200000, 3800000, 2900000, 1100000, 500000]
    region_share = [
        region_rev / sum(region_revenue) * 100 for region_rev in region_revenue
    ]

    geo_data = pd.DataFrame(
        {"Region": regions, "Revenue": region_revenue, "Share": region_share}
    )

    # Create the presentation
    prs = PPTXPresentation()

    # 1. Title slide
    title_slide = prs.add_slide(layout_type=0, title="FY 2024 Financial Performance")

    # Add subtitle and author
    title_slide.add_text_box(
        "Confidential Financial Report",
        position=(1, 3.5),
        size=(8, 1),
        align="center",
        font_size=18,
        bold=True,
    )

    title_slide.add_text_box(
        "Prepared by Finance Department\nMarch 09, 2025",
        position=(1, 5),
        size=(8, 1),
        align="center",
        font_size=14,
    )

    # 2. Executive Summary slide
    summary_slide = prs.add_slide(title="Executive Summary")

    # Add summary metrics
    metrics = [
        {
            "label": "Annual Revenue",
            "value": f"${sum(revenue_data)/1000000:.1f}M",
            "color": "4472C4",
        },
        {
            "label": "Annual Profit",
            "value": f"${sum(profit_data)/1000000:.1f}M",
            "color": "70AD47",
        },
        {
            "label": "Avg. Margin",
            "value": f"{sum(profit_data)/sum(revenue_data)*100:.1f}%",
            "color": "ED7D31",
        },
        {
            "label": "YoY Growth",
            "value": f"{np.mean(quarterly_data['YoY_Growth']):.1f}%",
            "color": "7030A0",
        },
    ]

    # Create visual metrics with shapes
    for i, metric in enumerate(metrics):
        # Position metrics in a 2x2 grid
        x = 1 + (i % 2) * 4.5
        y = 1.5 + (i // 2) * 2.5

        # Add background shape
        title_slide.add_shape(
            "rectangle",
            position=(x, y),
            size=(3.5, 1.8),
            fill_color=metric["color"],
            line_color=metric["color"],
            line_width=0,
        )

        # Add value text (large)
        title_slide.add_text_box(
            metric["value"],
            position=(x, y + 0.3),
            size=(3.5, 0.8),
            align="center",
            font_size=28,
            bold=True,
            color="FFFFFF",
        )

        # Add label text (smaller)
        title_slide.add_text_box(
            metric["label"],
            position=(x, y + 1.2),
            size=(3.5, 0.4),
            align="center",
            font_size=16,
            color="FFFFFF",
        )

    # Add summary bullet points
    summary_points = [
        "Revenue exceeded target by 7.2% with strong Q4 performance",
        "Profit margins improved in second half of the year",
        "Marketing efficiency increased with 15% higher ROI",
        "Product mix shift towards higher-margin offerings",
    ]

    summary_text = "Key Highlights:\n" + "\n".join(
        [f"• {point}" for point in summary_points]
    )

    summary_slide.add_text_box(
        summary_text, position=(1, 6), size=(8, 2), font_size=14, font_name="Calibri"
    )

    # 3. Quarterly Performance slide
    quarterly_slide = prs.add_slide(title="Quarterly Performance")

    # Add quarterly chart - clustered bar chart
    quarterly_slide.add_chart(
        chart_type="clustered_bar",
        data=quarterly_data,
        position=(0.5, 1.2),
        size=(5, 3.5),
        title="Quarterly Revenue and Profit",
        category_column="Quarter",
        series_columns=["Revenue", "Profit"],
        series_1_name="Revenue",
        series_2_name="Profit",
        series_1_color="4472C4",  # Blue
        series_2_color="70AD47",  # Green
        has_data_labels=True,
        data_label_number_format="$#,##0,,M",
    )

    # Add quarterly growth chart - line chart
    quarterly_slide.add_chart(
        chart_type="line",
        data=quarterly_data,
        position=(6, 1.2),
        size=(4, 3.5),
        title="YoY Growth Rate",
        category_column="Quarter",
        series_columns=["YoY_Growth"],
        series_1_name="Growth",
        series_1_color="ED7D31",  # Orange
        series_1_line_width=3,
        show_markers=True,
        has_data_labels=True,
        data_label_number_format="0.0%",
        value_axis_number_format="0.0%",
    )

    # Add summary table
    quarterly_slide.add_table(
        data=quarterly_data,
        position=(0.5, 5),
        table_width=9.5,
        col_widths=[1.5, 2.5, 2.5, 2, 1],
        column_formats={
            "Quarter": "text",
            "Revenue": "dollars",
            "Costs": "dollars",
            "Profit": "dollars",
            "YoY_Growth": "percentage",
        },
        has_header=True,
        header_style={"fill_color": "4472C4", "font_color": "FFFFFF", "font_size": 12},
    )

    # Apply conditional formatting to the table
    quarterly_slide.apply_conditional_formatting(
        quarterly_slide.tables[-1],
        rules=[
            {
                "type": "color_scale",
                "column": "Profit",
                "min_color": "F8696B",  # Red
                "max_color": "63BE7B",  # Green
            },
            {
                "type": "highlight_cells",
                "column": "YoY_Growth",
                "operator": "greater_than",
                "value": 15,
                "color": "C6EFCE",  # Light green
            },
        ],
        start_row=1,
    )

    # 4. Monthly Trend Analysis slide
    monthly_slide = prs.add_slide(title="Monthly Trend Analysis")

    # Add monthly revenue and profit chart
    monthly_slide.add_chart(
        chart_type="line",
        data=monthly_data,
        position=(0.5, 1.2),
        size=(9.5, 3.5),
        title="Monthly Revenue and Profit Trends",
        category_column="Month",
        series_columns=["Revenue", "Profit"],
        series_1_name="Revenue",
        series_2_name="Profit",
        series_1_color="4472C4",  # Blue
        series_2_color="70AD47",  # Green
        series_1_line_width=2.5,
        series_2_line_width=2.5,
        show_markers=True,
        value_axis_visible=True,
        value_axis_has_gridlines=True,
        data_label_number_format="$#,##0,K",
    )

    # Add margin % chart (secondary axis not directly supported, so separate chart)
    monthly_slide.add_chart(
        chart_type="line",
        data=monthly_data,
        position=(0.5, 5),
        size=(9.5, 2),
        title="Monthly Profit Margin %",
        category_column="Month",
        series_columns=["Margin"],
        series_1_name="Profit Margin",
        series_1_color="ED7D31",  # Orange
        series_1_line_width=2,
        show_markers=True,
        value_axis_visible=True,
        value_axis_number_format="0.0%",
        data_label_number_format="0.0%",
    )

    # 5. Expense Breakdown slide
    expense_slide = prs.add_slide(title="Expense Breakdown")

    # Add donut chart for expense categories
    expense_slide.add_chart(
        chart_type="donut",
        data=expense_data,
        position=(0.5, 1.2),
        size=(4.5, 4),
        title="Expense Distribution",
        category_column="Category",
        value_column="Amount",
        series_name="Expenses",
        has_legend=True,
        legend_position="right",
        segment_colors={
            "COGS": "4472C4",  # Blue
            "Marketing": "ED7D31",  # Orange
            "R&D": "70AD47",  # Green
            "Administration": "FFC000",  # Yellow
            "Other": "5B9BD5",  # Light blue
        },
        data_label_number_format="0.0%",
    )

    # Create a stacked bar chart showing monthly expense breakdown
    expense_monthly = pd.DataFrame(
        {
            "Month": months,
            "COGS": cost_data,
            "Marketing": marketing_data,
            "R&D": rd_data,
            "Admin": admin_data,
            "Other": [
                rev * 0.02 for rev in revenue_data
            ],  # Other expenses (2% of revenue)
        }
    )

    expense_slide.add_chart(
        chart_type="stacked_bar",
        data=expense_monthly,
        position=(5.5, 1.2),
        size=(5, 4),
        title="Monthly Expenses by Category",
        category_column="Month",
        series_columns=["COGS", "Marketing", "R&D", "Admin", "Other"],
        series_1_color="4472C4",  # Blue
        series_2_color="ED7D31",  # Orange
        series_3_color="70AD47",  # Green
        series_4_color="FFC000",  # Yellow
        series_5_color="5B9BD5",  # Light blue
        data_label_position="inside_end",
        data_label_font_color="FFFFFF",
        data_label_number_format="$#,##0,K",
    )

    # Add expense summary table
    total_expenses = sum(expense_values)
    expense_summary = pd.DataFrame(
        {
            "Category": expense_categories,
            "Amount": expense_values,
            "Percentage": [value / total_expenses * 100 for value in expense_values],
        }
    )

    expense_slide.add_table(
        data=expense_summary,
        position=(0.5, 5.5),
        table_width=9.5,
        col_widths=[3, 3.5, 3],
        column_formats={
            "Category": "text",
            "Amount": "dollars",
            "Percentage": "percentage",
        },
        has_header=True,
        header_style={"fill_color": "4472C4", "font_color": "FFFFFF", "font_size": 12},
    )

    # 6. Product Performance slide
    product_slide = prs.add_slide(title="Product Performance")

    # Add horizontal bar chart showing product revenue
    product_slide.add_chart(
        chart_type="clustered_bar",
        data=product_data,
        position=(0.5, 1.2),
        size=(9.5, 3),
        title="Product Revenue",
        category_column="Product",
        series_columns=["Revenue"],
        series_1_name="Revenue",
        series_1_color="4472C4",
        chart_type="bar_clustered",  # Use horizontal bars
        has_data_labels=True,
        data_label_number_format="$#,##0,,M",
    )

    # Add product metrics table
    product_slide.add_table(
        data=product_data,
        position=(0.5, 4.5),
        table_width=9.5,
        col_widths=[2.5, 3, 2, 2],
        column_formats={
            "Product": "text",
            "Revenue": "dollars",
            "Growth": "percentage",
            "Margin": "percentage",
        },
        has_header=True,
        header_style={"fill_color": "4472C4", "font_color": "FFFFFF", "font_size": 12},
    )

    # Apply conditional formatting to the product table
    product_slide.apply_conditional_formatting(
        product_slide.tables[-1],
        rules=[
            {
                "type": "highlight_cells",
                "column": "Growth",
                "operator": "less_than",
                "value": 0,
                "color": "FFCCCC",  # Light red
            },
            {
                "type": "color_scale",
                "column": "Margin",
                "min_color": "FFFFFF",  # White
                "max_color": "63BE7B",  # Green
            },
            {
                "type": "top_bottom",
                "column": "Revenue",
                "top": True,
                "rank": 2,
                "color": "D8E4BC",  # Light green
            },
        ],
        start_row=1,
    )

    # 7. Geographic Performance slide
    geo_slide = prs.add_slide(title="Geographic Performance")

    # Add donut chart for geographic revenue share
    geo_slide.add_chart(
        chart_type="donut",
        data=geo_data,
        position=(0.5, 1.2),
        size=(4.5, 4),
        title="Revenue by Region",
        category_column="Region",
        value_column="Revenue",
        series_name="Revenue",
        has_legend=True,
        legend_position="right",
        segment_colors={
            "North America": "4472C4",  # Blue
            "Europe": "70AD47",  # Green
            "Asia Pacific": "ED7D31",  # Orange
            "Latin America": "FFC000",  # Yellow
            "MEA": "5B9BD5",  # Light blue
        },
        data_label_number_format="0.0%",
    )

    # Add bar chart for absolute values
    geo_slide.add_chart(
        chart_type="bar",
        data=geo_data,
        position=(5.5, 1.2),
        size=(5, 4),
        title="Regional Revenue ($M)",
        category_column="Region",
        series_columns=["Revenue"],
        series_1_name="Revenue",
        series_1_color="4472C4",
        has_data_labels=True,
        data_label_number_format="$#,##0,,M",
    )

    # Add geographic table
    geo_slide.add_table(
        data=geo_data,
        position=(0.5, 5.5),
        table_width=9.5,
        col_widths=[3.5, 3, 3],
        column_formats={"Region": "text", "Revenue": "dollars", "Share": "percentage"},
        has_header=True,
        header_style={"fill_color": "4472C4", "font_color": "FFFFFF", "font_size": 12},
    )

    # 8. Conclusion slide
    conclusion_slide = prs.add_slide(title="Key Takeaways and Next Steps")

    # Left side: Key takeaways
    takeaways = [
        "Revenue growth exceeded expectations with 15.3% YoY increase",
        "Profit margins improved by 2.5 percentage points",
        "Product C performance needs attention",
        "Asia Pacific region continues to show accelerating growth",
        "Cost management initiatives yielded 8% operational efficiency",
    ]

    takeaways_text = "Key Takeaways:\n" + "\n".join(
        [f"• {point}" for point in takeaways]
    )

    conclusion_slide.add_text_box(
        takeaways_text,
        position=(0.5, 1.2),
        size=(4.5, 3),
        font_size=14,
        font_name="Calibri",
    )

    # Right side: Next steps
    next_steps = [
        "Increase investment in high-margin Product A and D",
        "Address Product C performance with targeted marketing",
        "Continue expansion in Asia Pacific with localized solutions",
        "Implement Phase 2 of cost optimization program",
        "Evaluate potential M&A opportunities in Q2",
    ]

    next_steps_text = "Next Steps:\n" + "\n".join(
        [f"• {point}" for point in next_steps]
    )

    conclusion_slide.add_text_box(
        next_steps_text,
        position=(5.5, 1.2),
        size=(4.5, 3),
        font_size=14,
        font_name="Calibri",
    )

    # Add footer
    conclusion_slide.add_text_box(
        "Confidential - For Internal Use Only",
        position=(0.5, 6.5),
        size=(9.5, 0.5),
        align="center",
        font_size=10,
        italic=True,
    )

    # Save the presentation
    prs.save("advanced_financial_report.pptx")
    print("Advanced financial report created: advanced_financial_report.pptx")


if __name__ == "__main__":
    create_financial_report()
