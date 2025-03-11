import pandas as pd
import numpy as np
from pptx_charts_tables import PPTXPresentation


def create_marketing_analysis():
    """
    Create a marketing campaign analysis presentation with detailed
    performance metrics, ROI visualization, and channel comparison.
    """
    # Create sample data for marketing campaign analysis
    campaigns = [
        "Spring Promotion",
        "Summer Sale",
        "Back to School",
        "Holiday Special",
        "Mobile App Launch",
    ]

    # Campaign performance data
    campaign_data = pd.DataFrame(
        {
            "Campaign": campaigns,
            "Budget": [125000, 200000, 150000, 300000, 175000],
            "Revenue": [420000, 680000, 510000, 1250000, 390000],
            "ROI": [236, 240, 240, 317, 123],
            "Leads": [5200, 7800, 6100, 15600, 4800],
            "Conversions": [780, 1360, 918, 2340, 720],
        }
    )

    # Calculate conversion rates
    campaign_data["Conversion_Rate"] = (
        campaign_data["Conversions"] / campaign_data["Leads"] * 100
    ).round(1)

    # Channel performance data
    channels = ["Social Media", "Email", "Search", "Display", "Affiliate", "Direct"]

    channel_data = pd.DataFrame(
        {
            "Channel": channels,
            "Impressions": [3500000, 1200000, 2100000, 4500000, 850000, 620000],
            "Clicks": [105000, 84000, 63000, 90000, 34000, 49600],
            "Cost": [85000, 45000, 120000, 150000, 51000, 30000],
            "Conversions": [2100, 3360, 1890, 1800, 680, 1488],
            "Revenue": [189000, 302400, 170100, 162000, 61200, 133920],
        }
    )

    # Calculate CTR, CPC, CVR, and ROAS
    channel_data["CTR"] = (
        channel_data["Clicks"] / channel_data["Impressions"] * 100
    ).round(2)
    channel_data["CPC"] = (channel_data["Cost"] / channel_data["Clicks"]).round(2)
    channel_data["CVR"] = (
        channel_data["Conversions"] / channel_data["Clicks"] * 100
    ).round(2)
    channel_data["ROAS"] = (channel_data["Revenue"] / channel_data["Cost"]).round(2)

    # Create audience demographics data
    age_groups = ["18-24", "25-34", "35-44", "45-54", "55-64", "65+"]

    demographics = pd.DataFrame(
        {
            "Age_Group": age_groups,
            "Audience_Size": [5600, 8200, 7400, 4300, 2800, 1700],
            "Conversion_Rate": [2.1, 3.8, 4.2, 3.6, 2.9, 1.5],
            "Avg_Order_Value": [72, 96, 118, 105, 88, 65],
        }
    )

    # Monthly performance for the year
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

    # Monthly marketing spend and revenue
    base_spend = 80000
    base_revenue = 240000
    seasonal_factor_spend = [0.7, 0.7, 0.9, 1.0, 1.3, 1.5, 1.2, 1.2, 1.0, 1.0, 1.5, 2.0]
    seasonal_factor_revenue = [
        0.65,
        0.7,
        0.85,
        1.05,
        1.3,
        1.6,
        1.15,
        1.2,
        1.05,
        1.1,
        1.6,
        2.5,
    ]

    monthly_data = pd.DataFrame(
        {
            "Month": months,
            "Spend": [base_spend * factor for factor in seasonal_factor_spend],
            "Revenue": [base_revenue * factor for factor in seasonal_factor_revenue],
        }
    )

    # Calculate ROI for each month
    monthly_data["ROI"] = (
        (monthly_data["Revenue"] - monthly_data["Spend"]) / monthly_data["Spend"] * 100
    ).round(1)

    # Campaign performance by week (for one selected campaign)
    weeks = ["Week 1", "Week 2", "Week 3", "Week 4", "Week 5", "Week 6"]

    weekly_data = pd.DataFrame(
        {
            "Week": weeks,
            "Impressions": [820000, 950000, 1100000, 1050000, 890000, 690000],
            "Clicks": [24600, 30400, 38500, 35700, 27600, 19800],
            "Conversions": [492, 760, 1155, 1071, 690, 396],
        }
    )

    weekly_data["CTR"] = (
        weekly_data["Clicks"] / weekly_data["Impressions"] * 100
    ).round(2)
    weekly_data["CVR"] = (
        weekly_data["Conversions"] / weekly_data["Clicks"] * 100
    ).round(2)

    # Create the presentation
    prs = PPTXPresentation()

    # 1. Title slide
    title_slide = prs.add_slide(layout_type=0, title="Marketing Campaign Analysis")

    # Add subtitle and date
    title_slide.add_text_box(
        "2024 Performance Report",
        position=(1, 3),
        size=(8, 0.8),
        align="center",
        font_size=20,
        bold=True,
    )

    title_slide.add_text_box(
        "Marketing Analytics Team\nMarch 10, 2025",
        position=(1, 4),
        size=(8, 0.8),
        align="center",
        font_size=16,
    )

    # 2. Campaign Overview slide
    campaign_slide = prs.add_slide(title="Campaign Performance Overview")

    # Add bar chart showing campaign ROI
    campaign_slide.add_chart(
        chart_type="bar",
        data=campaign_data,
        position=(0.5, 1.2),
        size=(4.5, 3),
        title="Campaign ROI (%)",
        category_column="Campaign",
        series_columns=["ROI"],
        series_1_name="ROI",
        series_1_color="70AD47",  # Green
        has_data_labels=True,
        data_label_number_format="0%",
    )

    # Add horizontal bar chart showing budget and revenue
    budget_revenue = campaign_data[["Campaign", "Budget", "Revenue"]]
    campaign_slide.add_chart(
        chart_type="clustered_bar",
        data=budget_revenue,
        position=(5.5, 1.2),
        size=(4.5, 3),
        title="Budget vs. Revenue",
        category_column="Campaign",
        series_columns=["Budget", "Revenue"],
        series_1_name="Budget",
        series_2_name="Revenue",
        series_1_color="5B9BD5",  # Blue
        series_2_color="ED7D31",  # Orange
        chart_type="bar_clustered",  # Horizontal bars
        has_data_labels=True,
        data_label_number_format="$#,##0,K",
    )

    # Add a table with campaign performance data
    campaign_slide.add_table(
        data=campaign_data,
        position=(0.5, 4.5),
        table_width=9.5,
        column_formats={
            "Campaign": "text",
            "Budget": "dollars",
            "Revenue": "dollars",
            "ROI": "percentage",
            "Leads": "counts",
            "Conversions": "counts",
            "Conversion_Rate": "percentage",
        },
        has_header=True,
        header_style={"fill_color": "70AD47", "font_color": "FFFFFF"},
    )

    # Apply conditional formatting to the table
    campaign_slide.apply_conditional_formatting(
        campaign_slide.tables[-1],
        rules=[
            {
                "type": "color_scale",
                "column": "ROI",
                "min_color": "FFFFFF",  # White
                "max_color": "63BE7B",  # Green
            },
            {
                "type": "highlight_cells",
                "column": "Conversion_Rate",
                "operator": "greater_than",
                "value": 15,
                "color": "D8E4BC",  # Light green
            },
        ],
        start_row=1,
    )

    # 3. Channel Analysis slide
    channel_slide = prs.add_slide(title="Marketing Channel Performance")

    # Add clustered bar chart showing conversions and ROAS by channel
    channel_metrics = channel_data[["Channel", "Conversions", "ROAS"]]

    # Create a copy with ROAS scaled for better visualization on the same chart
    channel_metrics_viz = channel_metrics.copy()
    channel_metrics_viz["ROAS_Scaled"] = (
        channel_metrics_viz["ROAS"] * 500
    )  # Scale for better visibility

    channel_slide.add_chart(
        chart_type="clustered_bar",
        data=channel_metrics_viz,
        position=(0.5, 1.2),
        size=(4.5, 3),
        title="Channel Performance: Conversions & ROAS",
        category_column="Channel",
        series_columns=["Conversions", "ROAS_Scaled"],
        series_1_name="Conversions",
        series_2_name="ROAS",
        series_1_color="5B9BD5",  # Blue
        series_2_color="ED7D31",  # Orange
        has_data_labels=True,
        data_label_number_format="0",
    )

    # Add donut chart showing channel revenue distribution
    channel_revenue = pd.DataFrame(
        {"Channel": channels, "Revenue": channel_data["Revenue"]}
    )

    channel_slide.add_chart(
        chart_type="donut",
        data=channel_revenue,
        position=(5.5, 1.2),
        size=(4, 3),
        title="Revenue by Channel",
        category_column="Channel",
        value_column="Revenue",
        series_name="Revenue",
        segment_colors={
            "Social Media": "5B9BD5",  # Blue
            "Email": "ED7D31",  # Orange
            "Search": "A5A5A5",  # Gray
            "Display": "70AD47",  # Green
            "Affiliate": "FFC000",  # Yellow
            "Direct": "4472C4",  # Dark Blue
        },
        has_legend=True,
        legend_position="right",
        data_label_number_format="0%",
    )

    # Add a table with channel performance data (selected metrics)
    channel_display = channel_data[
        ["Channel", "Impressions", "Clicks", "CTR", "CPC", "CVR", "ROAS"]
    ]
    channel_slide.add_table(
        data=channel_display,
        position=(0.5, 4.5),
        table_width=9.5,
        column_formats={
            "Channel": "text",
            "Impressions": "counts",
            "Clicks": "counts",
            "CTR": "percentage",
            "CPC": "dollars",
            "CVR": "percentage",
            "ROAS": "counts",
        },
        has_header=True,
        header_style={"fill_color": "5B9BD5", "font_color": "FFFFFF"},
    )

    # Apply conditional formatting to highlight best performing channels
    channel_slide.apply_conditional_formatting(
        channel_slide.tables[-1],
        rules=[
            {
                "type": "top_bottom",
                "column": "CTR",
                "top": True,
                "rank": 2,
                "color": "D8E4BC",  # Light green
            },
            {
                "type": "top_bottom",
                "column": "ROAS",
                "top": True,
                "rank": 2,
                "color": "D8E4BC",  # Light green
            },
        ],
        start_row=1,
    )

    # 4. Audience Demographics slide
    demo_slide = prs.add_slide(title="Audience Demographics Analysis")

    # Add bar chart showing audience size by age group
    demo_slide.add_chart(
        chart_type="bar",
        data=demographics,
        position=(0.5, 1.2),
        size=(4.5, 2.5),
        title="Audience Size by Age Group",
        category_column="Age_Group",
        series_columns=["Audience_Size"],
        series_1_name="Audience Size",
        series_1_color="4472C4",  # Blue
        has_data_labels=True,
        data_label_number_format="0,K",
    )

    # Add line chart showing conversion rate by age group
    demo_slide.add_chart(
        chart_type="line",
        data=demographics,
        position=(5.5, 1.2),
        size=(4.5, 2.5),
        title="Conversion Rate by Age Group",
        category_column="Age_Group",
        series_columns=["Conversion_Rate"],
        series_1_name="Conversion Rate",
        series_1_color="ED7D31",  # Orange
        series_1_line_width=2.5,
        show_markers=True,
        value_axis_visible=True,
        value_axis_number_format="0.0%",
        has_data_labels=True,
        data_label_number_format="0.0%",
    )

    # Add bar chart showing average order value by age group
    demo_slide.add_chart(
        chart_type="bar",
        data=demographics,
        position=(0.5, 4),
        size=(4.5, 2.5),
        title="Average Order Value by Age Group",
        category_column="Age_Group",
        series_columns=["Avg_Order_Value"],
        series_1_name="Avg Order Value",
        series_1_color="70AD47",  # Green
        has_data_labels=True,
        data_label_number_format="$0",
    )

    # Add text box with demographics insights
    demo_insights = """Key Demographics Insights:
• 25-44 age group represents our largest audience segment (53%)
• 35-44 has highest conversion rate (4.2%) and average order value ($118)
• 65+ shows lowest performance across all metrics
• 25-34 segment has grown 18% this year
• Mobile usage highest in 18-34 demographic (76%)"""

    demo_slide.add_text_box(
        demo_insights, position=(5.5, 4), size=(4.5, 2.5), font_size=14, no_fill=True
    )

    # 5. Monthly Trends slide
    monthly_slide = prs.add_slide(title="Monthly Marketing Performance")

    # Add line chart showing monthly spend and revenue
    monthly_slide.add_chart(
        chart_type="line",
        data=monthly_data,
        position=(0.5, 1.2),
        size=(9, 3),
        title="Monthly Marketing Spend vs. Revenue",
        category_column="Month",
        series_columns=["Spend", "Revenue"],
        series_1_name="Marketing Spend",
        series_2_name="Revenue",
        series_1_color="5B9BD5",  # Blue
        series_2_color="ED7D31",  # Orange
        series_1_line_width=2,
        series_2_line_width=2,
        show_markers=True,
        value_axis_visible=True,
        value_axis_has_gridlines=True,
        data_label_number_format="$#,##0,K",
    )

    # Add bar chart showing monthly ROI
    monthly_slide.add_chart(
        chart_type="bar",
        data=monthly_data,
        position=(0.5, 4.5),
        size=(9, 2),
        title="Monthly Marketing ROI (%)",
        category_column="Month",
        series_columns=["ROI"],
        series_1_name="ROI",
        series_1_color="70AD47",  # Green
        has_data_labels=True,
        data_label_number_format="0.0%",
    )

    # 6. Weekly Campaign Performance slide
    weekly_slide = prs.add_slide(title="Weekly Campaign Performance (Holiday Special)")

    # Add line chart showing weekly impressions and clicks
    weekly_metrics = weekly_data[["Week", "Impressions", "Clicks"]]

    # Scale clicks for better visualization on the same chart
    weekly_metrics["Clicks_Scaled"] = weekly_metrics["Clicks"] * 10

    weekly_slide.add_chart(
        chart_type="line",
        data=weekly_metrics,
        position=(0.5, 1.2),
        size=(4.5, 2.5),
        title="Weekly Impressions & Clicks",
        category_column="Week",
        series_columns=["Impressions", "Clicks_Scaled"],
        series_1_name="Impressions",
        series_2_name="Clicks (×10)",
        series_1_color="5B9BD5",  # Blue
        series_2_color="ED7D31",  # Orange
        series_1_line_width=2,
        series_2_line_width=2,
        show_markers=True,
        value_axis_visible=True,
        data_label_number_format="#,##0,K",
    )

    # Add line chart showing CTR and CVR trends
    conversion_metrics = weekly_data[["Week", "CTR", "CVR"]]
    weekly_slide.add_chart(
        chart_type="line",
        data=conversion_metrics,
        position=(5.5, 1.2),
        size=(4.5, 2.5),
        title="Weekly CTR & CVR Trends",
        category_column="Week",
        series_columns=["CTR", "CVR"],
        series_1_name="Click-Through Rate",
        series_2_name="Conversion Rate",
        series_1_color="70AD47",  # Green
        series_2_color="FFC000",  # Yellow
        series_1_line_width=2,
        series_2_line_width=2,
        show_markers=True,
        value_axis_visible=True,
        value_axis_number_format="0.0%",
        data_label_number_format="0.00%",
    )

    # Add a table with weekly performance data
    weekly_slide.add_table(
        data=weekly_data,
        position=(0.5, 4),
        table_width=9.5,
        column_formats={
            "Week": "text",
            "Impressions": "counts",
            "Clicks": "counts",
            "Conversions": "counts",
            "CTR": "percentage",
            "CVR": "percentage",
        },
        has_header=True,
        header_style={"fill_color": "70AD47", "font_color": "FFFFFF"},
    )

    # Apply conditional formatting to highlight best and worst weeks
    weekly_slide.apply_conditional_formatting(
        weekly_slide.tables[-1],
        rules=[
            {
                "type": "color_scale",
                "column": "CVR",
                "min_color": "FFCCCC",  # Light red
                "max_color": "D8E4BC",  # Light green
            }
        ],
        start_row=1,
    )

    # 7. Conclusions & Recommendations slide
    conclusion_slide = prs.add_slide(title="Conclusions & Recommendations")

    # Key findings
    findings = [
        "Holiday Special campaign delivered highest ROI (317%)",
        "Email channel shows strongest overall performance (ROAS 6.7)",
        "35-44 age demographic remains most valuable segment",
        "November and December show highest marketing efficiency",
        "Week 3 of campaigns typically shows peak performance",
    ]

    # Recommendations
    recommendations = [
        "Increase budget allocation to Email channel by 20%",
        "Optimize Mobile App campaign for higher conversion rate",
        "Develop targeted content for 35-44 age demographic",
        "Extend Holiday campaign duration for next year",
        "Implement A/B testing for Social Media creatives",
    ]

    # Left side: Key findings
    findings_text = "Key Findings:\n" + "\n".join([f"• {point}" for point in findings])
    conclusion_slide.add_text_box(
        findings_text, position=(0.5, 1.2), size=(4.5, 3), font_size=14, no_fill=True
    )

    # Right side: Recommendations
    recommendations_text = "Recommendations:\n" + "\n".join(
        [f"• {point}" for point in recommendations]
    )
    conclusion_slide.add_text_box(
        recommendations_text,
        position=(5.5, 1.2),
        size=(4.5, 3),
        font_size=14,
        no_fill=True,
    )

    # Add a shape highlighting projected impact
    conclusion_slide.add_shape(
        "rectangle",
        position=(2.5, 4.5),
        size=(5, 1.5),
        fill_color="70AD47",  # Green
        text="Projected Impact:\nImplementing these recommendations is estimated to\nincrease overall marketing ROI by 18-24% in 2025",
        font_size=16,
        font_color="FFFFFF",
        align="center",
        v_align="middle",
    )

    # Add a footer
    conclusion_slide.add_text_box(
        "Marketing Analytics Team - Confidential",
        position=(0.5, 6.5),
        size=(9.5, 0.3),
        align="center",
        font_size=8,
        italic=True,
    )

    # Save the presentation
    prs.save("marketing_analysis.pptx")
    print("Marketing Analysis presentation created: marketing_analysis.pptx")


if __name__ == "__main__":
    create_marketing_analysis()
