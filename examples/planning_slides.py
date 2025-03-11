import pandas as pd
import numpy as np
from pptx_charts_tables import PPTXPresentation


def create_budget_presentation():
    """
    Create a budget planning presentation with conditional formatting,
    stacked charts, and variance analysis.
    """
    # Create sample data
    departments = ["Marketing", "Sales", "IT", "Operations", "HR", "R&D", "Finance"]

    # Budget data
    last_year = [1250000, 1850000, 980000, 1650000, 520000, 1150000, 410000]
    this_year = [1350000, 2050000, 1150000, 1550000, 540000, 1450000, 430000]
    change_pct = [(t - l) / l * 100 for t, l in zip(this_year, last_year)]

    budget_data = pd.DataFrame(
        {
            "Department": departments,
            "Last_Year": last_year,
            "This_Year": this_year,
            "Change_Pct": change_pct,
        }
    )

    # Quarterly budget breakdown
    quarters = ["Q1", "Q2", "Q3", "Q4"]
    seasonal_weights = {
        "Marketing": [0.22, 0.26, 0.28, 0.24],
        "Sales": [0.20, 0.25, 0.25, 0.30],
        "IT": [0.25, 0.25, 0.25, 0.25],
        "Operations": [0.23, 0.24, 0.26, 0.27],
        "HR": [0.28, 0.24, 0.24, 0.24],
        "R&D": [0.20, 0.30, 0.30, 0.20],
        "Finance": [0.30, 0.20, 0.20, 0.30],
    }

    # Create quarterly data for each department
    quarterly_data = {}
    for dept in departments:
        quarterly_budget = [
            this_year[departments.index(dept)] * w for w in seasonal_weights[dept]
        ]
        quarterly_data[dept] = quarterly_budget

    # Create expense categories
    expense_categories = ["Personnel", "Equipment", "Services", "Facilities", "Other"]

    # Generate expense breakdowns for each department
    expense_breakdown = pd.DataFrame(
        {
            "Department": departments,
            "Personnel": [850000, 1350000, 750000, 800000, 420000, 1050000, 320000],
            "Equipment": [120000, 250000, 250000, 350000, 30000, 200000, 25000],
            "Services": [180000, 200000, 100000, 150000, 50000, 100000, 45000],
            "Facilities": [150000, 180000, 30000, 200000, 25000, 70000, 25000],
            "Other": [50000, 70000, 20000, 50000, 15000, 30000, 15000],
        }
    )

    # Calculate totals to ensure they match this_year
    expense_breakdown["Total"] = expense_breakdown[expense_categories].sum(axis=1)

    # Budget variance (actual vs planned for current year to date)
    planned = [
        900000,
        1366667,
        766667,
        1033333,
        360000,
        966667,
        286667,
    ]  # 2/3 of year plan
    actual = [950000, 1400000, 800000, 1050000, 320000, 1000000, 300000]
    variance = [a - p for a, p in zip(actual, planned)]
    variance_pct = [v / p * 100 for v, p in zip(variance, planned)]

    variance_data = pd.DataFrame(
        {
            "Department": departments,
            "Planned": planned,
            "Actual": actual,
            "Variance": variance,
            "Variance_Pct": variance_pct,
        }
    )

    # Create the presentation
    prs = PPTXPresentation()

    # 1. Title slide
    title_slide = prs.add_slide(layout_type=0, title="FY 2025 Budget Planning")

    # Add subtitle and date
    title_slide.add_text_box(
        "Budget Review and Planning",
        position=(1, 3),
        size=(8, 0.8),
        align="center",
        font_size=20,
        bold=True,
    )

    title_slide.add_text_box(
        "Finance Department\nMarch 10, 2025",
        position=(1, 4),
        size=(8, 0.8),
        align="center",
        font_size=16,
    )

    # 2. Budget Overview slide
    overview_slide = prs.add_slide(title="Budget Overview by Department")

    # Add clustered bar chart comparing last year vs this year
    overview_slide.add_chart(
        chart_type="clustered_bar",
        data=budget_data,
        position=(0.5, 1.2),
        size=(5, 3.5),
        title="Budget Comparison: Last Year vs This Year",
        category_column="Department",
        series_columns=["Last_Year", "This_Year"],
        series_1_name="FY 2024",
        series_2_name="FY 2025",
        series_1_color="8497B0",  # Gray-blue
        series_2_color="5B9BD5",  # Blue
        has_data_labels=True,
        data_label_number_format="$#,##0,K",
    )

    # Add chart showing percentage change
    overview_slide.add_chart(
        chart_type="bar",
        data=budget_data,
        position=(6, 1.2),
        size=(4, 3.5),
        title="Budget Change (%)",
        category_column="Department",
        series_columns=["Change_Pct"],
        series_1_name="% Change",
        series_1_color="70AD47",  # Green
        has_data_labels=True,
        data_label_number_format="0.0%",
    )

    # Add a table with the budget data and formatting
    overview_slide.add_table(
        data=budget_data,
        position=(0.5, 5),
        table_width=9.5,
        column_formats={
            "Department": "text",
            "Last_Year": "dollars",
            "This_Year": "dollars",
            "Change_Pct": "percentage",
        },
        has_header=True,
        header_style={"fill_color": "5B9BD5", "font_color": "FFFFFF"},
    )

    # Apply conditional formatting to the table
    overview_slide.apply_conditional_formatting(
        overview_slide.tables[-1],
        rules=[
            {
                "type": "color_scale",
                "column": "Change_Pct",
                "min_color": "F8696B",  # Red (for negative)
                "mid_color": "FFEB84",  # Yellow (for around 0)
                "max_color": "63BE7B",  # Green (for positive)
            }
        ],
        start_row=1,
    )

    # 3. Quarterly Breakdown slide
    # First, prepare data in the right format for a stacked bar chart
    q1_data = [quarterly_data[dept][0] for dept in departments]
    q2_data = [quarterly_data[dept][1] for dept in departments]
    q3_data = [quarterly_data[dept][2] for dept in departments]
    q4_data = [quarterly_data[dept][3] for dept in departments]

    quarterly_df = pd.DataFrame(
        {
            "Department": departments,
            "Q1": q1_data,
            "Q2": q2_data,
            "Q3": q3_data,
            "Q4": q4_data,
        }
    )

    quarterly_slide = prs.add_slide(title="Quarterly Budget Breakdown")

    # Add stacked bar chart showing quarterly breakdown
    quarterly_slide.add_chart(
        chart_type="stacked_bar",
        data=quarterly_df,
        position=(0.5, 1.2),
        size=(9, 3.5),
        title="Budget Allocation by Quarter",
        category_column="Department",
        series_columns=["Q1", "Q2", "Q3", "Q4"],
        series_1_color="5B9BD5",  # Blue
        series_2_color="ED7D31",  # Orange
        series_3_color="A5A5A5",  # Gray
        series_4_color="70AD47",  # Green
        data_label_position="inside_end",
        data_label_font_color="FFFFFF",
        data_label_number_format="$#,##0,K",
    )

    # Add table with quarterly data
    quarterly_slide.add_table(
        data=quarterly_df,
        position=(0.5, 5),
        table_width=9.5,
        column_formats={
            "Department": "text",
            "Q1": "dollars",
            "Q2": "dollars",
            "Q3": "dollars",
            "Q4": "dollars",
        },
        has_header=True,
        header_style={"fill_color": "5B9BD5", "font_color": "FFFFFF"},
    )

    # Add subtle highlighting to Q4 column (which has highest budget for many departments)
    quarterly_slide.apply_conditional_formatting(
        quarterly_slide.tables[-1],
        rules=[
            {
                "type": "highlight_cells",
                "column": "Q4",
                "operator": "greater_than",
                "value": 350000,
                "color": "E6F2FF",  # Light blue
            }
        ],
        start_row=1,
    )

    # 4. Expense Categories slide
    expense_slide = prs.add_slide(title="Budget Breakdown by Expense Category")

    # Select a subset of departments for clarity
    selected_depts = ["Marketing", "Sales", "IT", "R&D"]
    selected_expense_data = expense_breakdown[
        expense_breakdown["Department"].isin(selected_depts)
    ]

    # Create a more presentation-friendly DataFrame without the Total column
    chart_expense_data = selected_expense_data[["Department"] + expense_categories]

    # Add stacked bar chart for expense categories
    expense_slide.add_chart(
        chart_type="stacked_bar",
        data=chart_expense_data,
        position=(0.5, 1.2),
        size=(9, 3),
        title="Expense Categories by Department",
        category_column="Department",
        series_columns=expense_categories,
        series_1_color="5B9BD5",  # Blue
        series_2_color="ED7D31",  # Orange
        series_3_color="A5A5A5",  # Gray
        series_4_color="70AD47",  # Green
        series_5_color="FFC000",  # Yellow
        data_label_position="inside_end",
        data_label_font_color="FFFFFF",
        data_label_number_format="$#,##0,K",
    )

    # Create a donut chart for overall expense breakdown (using first department as example)
    dept_example = "Sales"
    dept_data = expense_breakdown[expense_breakdown["Department"] == dept_example]

    # Transform to format needed for donut chart
    dept_expenses = pd.DataFrame(
        {
            "Category": expense_categories,
            "Amount": dept_data[expense_categories].values[0],
        }
    )

    # Add donut chart
    expense_slide.add_chart(
        chart_type="donut",
        data=dept_expenses,
        position=(0.5, 4.5),
        size=(4, 2.5),
        title=f"{dept_example} Department Expenses",
        category_column="Category",
        value_column="Amount",
        series_name="Expenses",
        segment_colors={
            "Personnel": "5B9BD5",  # Blue
            "Equipment": "ED7D31",  # Orange
            "Services": "A5A5A5",  # Gray
            "Facilities": "70AD47",  # Green
            "Other": "FFC000",  # Yellow
        },
        has_legend=True,
        legend_position="right",
        data_label_number_format="0%",
    )

    # Add a text box with expense breakdown insights
    expense_insights = f"""Key Expense Insights:
• Personnel costs represent the largest expense category across all departments
• IT has the highest proportion of equipment expenses (22%)
• Sales has significant service costs for customer relationship tools
• R&D investment increased by 26% this year
• Personnel costs average 70% of total budget"""

    expense_slide.add_text_box(
        expense_insights, position=(5, 4.5), size=(5, 2.5), font_size=14, no_fill=True
    )

    # 5. Variance Analysis slide
    variance_slide = prs.add_slide(title="Budget Variance Analysis (YTD)")

    # Add clustered bar chart for variance
    variance_slide.add_chart(
        chart_type="clustered_bar",
        data=variance_data,
        position=(0.5, 1.2),
        size=(5, 3.5),
        title="Planned vs Actual Budget (YTD)",
        category_column="Department",
        series_columns=["Planned", "Actual"],
        series_1_name="Planned",
        series_2_name="Actual",
        series_1_color="5B9BD5",  # Blue
        series_2_color="ED7D31",  # Orange
        has_data_labels=True,
        data_label_number_format="$#,##0,K",
    )

    # Add bar chart showing variance percentage
    variance_slide.add_chart(
        chart_type="bar",
        data=variance_data,
        position=(6, 1.2),
        size=(4, 3.5),
        title="Budget Variance (%)",
        category_column="Department",
        series_columns=["Variance_Pct"],
        series_1_name="% Variance",
        series_1_color="70AD47",  # Green
        has_data_labels=True,
        data_label_number_format="+0.0%",
    )

    # Add table with variance data
    variance_slide.add_table(
        data=variance_data,
        position=(0.5, 5),
        table_width=9.5,
        column_formats={
            "Department": "text",
            "Planned": "dollars",
            "Actual": "dollars",
            "Variance": "dollars",
            "Variance_Pct": "percentage",
        },
        has_header=True,
        header_style={"fill_color": "5B9BD5", "font_color": "FFFFFF"},
    )

    # Apply conditional formatting to highlight positive and negative variances
    variance_slide.apply_conditional_formatting(
        variance_slide.tables[-1],
        rules=[
            {
                "type": "highlight_cells",
                "column": "Variance",
                "operator": "less_than",
                "value": 0,
                "color": "FFCCCC",  # Light red
            },
            {
                "type": "highlight_cells",
                "column": "Variance",
                "operator": "greater_than",
                "value": 20000,
                "color": "E2EFDA",  # Light green
            },
        ],
        start_row=1,
    )

    # 6. Conclusion & Next Steps slide
    conclusion_slide = prs.add_slide(title="Conclusions & Next Steps")

    # Budget conclusions
    conclusions = [
        "Total budget increased by 12.6% compared to previous fiscal year",
        "IT and R&D departments show highest budget growth rates",
        "Q4 has highest overall budget allocation across departments",
        "Most departments are tracking within 5% of planned year-to-date spending",
        "Personnel costs remain the dominant expense category",
    ]

    # Next steps
    next_steps = [
        "Review Q3 spending plans for IT department",
        "Approve R&D budget increase for new product development",
        "Reassess HR department underspending",
        "Prepare mid-year review presentation for executive team",
        "Begin preliminary planning for FY 2026 budget cycle",
    ]

    # Add conclusions
    conclusions_text = "Budget Conclusions:\n" + "\n".join(
        [f"• {point}" for point in conclusions]
    )
    conclusion_slide.add_text_box(
        conclusions_text, position=(0.5, 1.2), size=(4.5, 3), font_size=14, no_fill=True
    )

    # Add next steps
    next_steps_text = "Next Steps:\n" + "\n".join(
        [f"• {point}" for point in next_steps]
    )
    conclusion_slide.add_text_box(
        next_steps_text, position=(5.5, 1.2), size=(4.5, 3), font_size=14, no_fill=True
    )

    # Add a timeline for budget review process
    timeline_months = ["March", "April", "May", "June", "July", "August"]
    timeline_activities = [
        "Mid-year review",
        "Dept. updates",
        "Strategic planning",
        "Preliminary estimates",
        "Budget workshops",
        "Final approvals",
    ]

    timeline_y = 4.5
    start_x = 0.5

    # Add the timeline header
    conclusion_slide.add_text_box(
        "Budget Review Timeline",
        position=(start_x, timeline_y - 0.5),
        size=(9.5, 0.4),
        align="center",
        font_size=14,
        bold=True,
        no_fill=True,
    )

    # Create timeline with shapes and connectors
    for i, (month, activity) in enumerate(zip(timeline_months, timeline_activities)):
        # Calculate position
        x = start_x + i * 1.5

        # Add month marker (circle)
        conclusion_slide.add_shape(
            "oval",
            position=(x, timeline_y),
            size=(0.5, 0.5),
            fill_color="5B9BD5",  # Blue
            line_color="5B9BD5",
        )

        # Add month label
        conclusion_slide.add_text_box(
            month,
            position=(x - 0.25, timeline_y + 0.6),
            size=(1, 0.3),
            align="center",
            font_size=10,
            bold=True,
            no_fill=True,
        )

        # Add activity label
        conclusion_slide.add_text_box(
            activity,
            position=(x - 0.5, timeline_y - 0.4),
            size=(1.5, 0.3),
            align="center",
            font_size=9,
            no_fill=True,
        )

        # Add connector between circles (except for the last one)
        if i < len(timeline_months) - 1:
            conclusion_slide.add_arrow(
                start_pos=(x + 0.5, timeline_y + 0.25),
                end_pos=(x + 1, timeline_y + 0.25),
                color="5B9BD5",
                width=1.5,
            )

    # Add a footer
    conclusion_slide.add_text_box(
        "Finance Department - Confidential",
        position=(0.5, 6.5),
        size=(9.5, 0.3),
        align="center",
        font_size=8,
        italic=True,
    )

    # Save the presentation
    prs.save("budget_planning.pptx")
    print("Budget Planning presentation created: budget_planning.pptx")


if __name__ == "__main__":
    create_budget_presentation()
