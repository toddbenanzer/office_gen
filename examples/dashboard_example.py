import pandas as pd
from pptx_charts_tables import PPTXPresentation


def create_sales_dashboard():
    """
    Create a sales dashboard presentation with multiple chart types.
    Shows regional sales performance with various visualizations.
    """
    # Create sample data
    regions = ["North", "South", "East", "West", "Central"]

    # Quarterly sales data by region
    quarterly_sales = pd.DataFrame(
        {
            "Region": regions,
            "Q1": [324000, 278000, 352000, 411000, 265000],
            "Q2": [368000, 301000, 389000, 425000, 292000],
            "Q3": [389000, 318000, 405000, 447000, 311000],
            "Q4": [452000, 365000, 438000, 510000, 348000],
        }
    )

    # Sales growth data
    growth_data = pd.DataFrame(
        {"Region": regions, "YoY_Growth": [12.8, 8.5, 15.2, 17.9, 9.4]}
    )

    # Product mix data
    product_categories = ["Electronics", "Apparel", "Home Goods", "Beauty", "Food"]
    product_mix = pd.DataFrame(
        {
            "Category": product_categories,
            "Sales": [1250000, 980000, 725000, 495000, 350000],
        }
    )

    # Monthly trend data
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

    # Base value with seasonal pattern
    base_sales = 800000
    seasonal_factor = [0.8, 0.85, 0.95, 1.0, 1.05, 1.1, 1.15, 1.2, 1.1, 1.05, 1.15, 1.4]

    monthly_trend = pd.DataFrame(
        {
            "Month": months,
            "Sales": [base_sales * factor for factor in seasonal_factor],
            "Target": [
                base_sales * factor * 1.1 for factor in seasonal_factor
            ],  # 10% higher target
        }
    )

    # Create the presentation
    prs = PPTXPresentation()

    # 1. Title slide
    title_slide = prs.add_slide(layout_type=0, title="Sales Performance Dashboard")
    title_slide.add_text_box(
        "Fiscal Year 2024",
        position=(1, 3.5),
        size=(8, 1),
        align="center",
        font_size=24,
        bold=True,
    )

    # 2. Regional Performance Overview slide
    regional_slide = prs.add_slide(title="Regional Sales Performance")

    # Add clustered bar chart showing quarterly performance by region
    regional_slide.add_chart(
        chart_type="clustered_bar",
        data=quarterly_sales,
        position=(0.5, 1.2),
        size=(9, 3.5),
        title="Quarterly Sales by Region",
        category_column="Region",
        series_columns=["Q1", "Q2", "Q3", "Q4"],
        series_1_color="4472C4",  # Blue
        series_2_color="ED7D31",  # Orange
        series_3_color="A5A5A5",  # Gray
        series_4_color="70AD47",  # Green
        has_data_labels=True,
        data_label_number_format="$#,##0,K",
    )

    # Add a table showing the same data
    regional_slide.add_table(
        data=quarterly_sales,
        position=(0.5, 5),
        table_width=9,
        has_header=True,
        column_formats={
            "Region": "text",
            "Q1": "dollars",
            "Q2": "dollars",
            "Q3": "dollars",
            "Q4": "dollars",
        },
        header_style={"fill_color": "4472C4", "font_color": "FFFFFF"},
    )

    # 3. Growth Analysis slide
    growth_slide = prs.add_slide(title="Regional Growth Analysis")

    # Add bar chart showing growth by region
    growth_slide.add_chart(
        chart_type="bar",
        data=growth_data,
        position=(0.5, 1.2),
        size=(4.5, 3.5),
        title="Year-over-Year Growth",
        category_column="Region",
        series_columns=["YoY_Growth"],
        series_1_color="70AD47",  # Green
        has_data_labels=True,
        data_label_number_format="0.0%",
        value_axis_number_format="0.0%",
    )

    # Add a shape highlighting top performer
    growth_slide.add_shape(
        "rectangle",
        position=(5.5, 1.5),
        size=(4, 1.5),
        fill_color="70AD47",  # Green
        text="West Region: Top Performer\n17.9% Growth",
        font_size=18,
        font_color="FFFFFF",
        align="center",
        v_align="middle",
    )

    # Add text box with analysis
    analysis_text = """Key Growth Insights:
• West region leads with 17.9% growth, driven by new store openings
• East shows strong performance at 15.2% growth
• All regions exceeded 8% growth target
• North recovered well from previous year's slump
• Central region needs additional marketing support"""

    growth_slide.add_text_box(
        analysis_text, position=(5.5, 3.5), size=(4, 2), font_size=14, no_fill=True
    )

    # 4. Product Mix slide
    product_slide = prs.add_slide(title="Product Category Analysis")

    # Add donut chart showing product mix
    product_slide.add_chart(
        chart_type="donut",
        data=product_mix,
        position=(0.5, 1.2),
        size=(4.5, 4),
        title="Sales by Product Category",
        category_column="Category",
        value_column="Sales",
        series_name="Sales",
        segment_colors={
            "Electronics": "4472C4",  # Blue
            "Apparel": "ED7D31",  # Orange
            "Home Goods": "FFC000",  # Yellow
            "Beauty": "70AD47",  # Green
            "Food": "5B9BD5",  # Light blue
        },
        has_legend=True,
        legend_position="right",
        data_label_number_format="0%",
    )

    # Add a table with the product mix data
    product_slide.add_table(
        data=product_mix,
        position=(5.5, 2),
        table_width=4,
        has_header=True,
        column_formats={"Category": "text", "Sales": "dollars"},
        header_style={"fill_color": "4472C4", "font_color": "FFFFFF"},
    )

    # Add text box with key insights
    product_insights = """Key Product Insights:
• Electronics remains our strongest category (33%)
• Apparel showing steady growth year-over-year
• Home Goods exceeded forecast by 12%
• Beauty products have highest profit margin (42%)
• Food category needs expanded selection"""

    product_slide.add_text_box(
        product_insights, position=(5.5, 4), size=(4, 2), font_size=14, no_fill=True
    )

    # 5. Monthly Trend slide
    trend_slide = prs.add_slide(title="Monthly Sales Trend")

    # Add line chart showing monthly trend
    trend_slide.add_chart(
        chart_type="line",
        data=monthly_trend,
        position=(0.5, 1.2),
        size=(9, 4),
        title="Monthly Sales vs Target",
        category_column="Month",
        series_columns=["Sales", "Target"],
        series_1_name="Actual Sales",
        series_2_name="Target",
        series_1_color="4472C4",  # Blue
        series_2_color="ED7D31",  # Orange
        series_1_line_width=2.5,
        series_2_line_width=1.5,
        series_2_line_style="dash",
        show_markers=True,
        value_axis_visible=True,
        value_axis_has_gridlines=True,
        data_label_number_format="$#,##0,K",
    )

    # Add a text box with insights
    trend_insights = """Monthly Performance Insights:
• Strong holiday season performance in Q4
• Exceeded targets in 7 out of 12 months
• Summer months show consistent growth
• February remains challenging due to seasonal factors
• December sales set new company record"""

    trend_slide.add_text_box(
        trend_insights, position=(0.5, 5.5), size=(9, 1.5), font_size=14, no_fill=True
    )

    # 6. Conclusion slide
    conclusion_slide = prs.add_slide(title="Summary and Next Steps")

    # Key takeaways in a structured format
    takeaways = [
        "Overall sales growth of 13.5% year-over-year",
        "West and East regions outperforming other areas",
        "Electronics and Apparel driving revenue growth",
        "Holiday season exceeded expectations",
        "Central region and Food category need attention",
    ]

    next_steps = [
        "Expand Electronics selection in all regions",
        "Launch Apparel marketing campaign in Central region",
        "Develop growth plan for Food category",
        "Increase inventory levels for Q4 2025",
        "Roll out staff training program in Q2",
    ]

    # Left side: Key takeaways
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

    # Add a footer
    conclusion_slide.add_text_box(
        "For more information, contact the Sales Analytics Team",
        position=(0.5, 6.5),
        size=(9.5, 0.5),
        align="center",
        font_size=10,
        italic=True,
    )

    # Save the presentation
    prs.save("sales_dashboard.pptx")
    print("Sales Dashboard created: sales_dashboard.pptx")


if __name__ == "__main__":
    create_sales_dashboard()
