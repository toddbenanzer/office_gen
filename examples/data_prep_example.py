import pandas as pd
import numpy as np
from datetime import datetime, timedelta
import calendar
from pptx_charts_tables import PPTXPresentation


def prepare_and_visualize_sales_data():
    """
    Example function that demonstrates how to:
    1. Create or load sample sales data
    2. Clean and transform the data
    3. Perform various aggregations and calculations
    4. Prepare different dataframes for specific chart types
    5. Create a PowerPoint presentation with the prepared data
    """
    # -----------------------------------------
    # Step 1: Create or load sample data
    # -----------------------------------------
    # In a real scenario, you might load from CSV/Excel:
    # raw_data = pd.read_csv('sales_data.csv')
    # raw_data = pd.read_excel('sales_data.xlsx')

    # For this example, we'll generate sample data
    raw_data = generate_sample_sales_data()

    print(f"Raw data shape: {raw_data.shape}")
    print("\nSample of raw data:")
    print(raw_data.head())

    # -----------------------------------------
    # Step 2: Clean and transform the data
    # -----------------------------------------
    # Handle missing values
    data = raw_data.copy()

    # Fill missing customer segments
    data["customer_segment"] = data["customer_segment"].fillna("Unknown")

    # Remove rows with missing transaction dates
    data = data.dropna(subset=["transaction_date"])

    # Fix data types
    data["transaction_date"] = pd.to_datetime(data["transaction_date"])
    data["quantity"] = data["quantity"].astype(int)

    # Create derived columns
    data["year"] = data["transaction_date"].dt.year
    data["month"] = data["transaction_date"].dt.month
    data["month_name"] = data["transaction_date"].dt.strftime("%b")
    data["quarter"] = data["transaction_date"].dt.quarter
    data["quarter_name"] = "Q" + data["quarter"].astype(str)
    data["day_of_week"] = data["transaction_date"].dt.day_name()

    # Calculate revenue
    data["revenue"] = data["quantity"] * data["unit_price"]

    # Calculate cost and profit
    data["cost"] = data["quantity"] * data["unit_cost"]
    data["profit"] = data["revenue"] - data["cost"]
    data["profit_margin"] = (data["profit"] / data["revenue"] * 100).round(2)

    print("\nData after cleaning and transformation:")
    print(data.head())

    # -----------------------------------------
    # Step 3: Perform various aggregations
    # -----------------------------------------

    # 3.1: Sales by product category
    product_sales = (
        data.groupby("product_category")
        .agg({"revenue": "sum", "profit": "sum", "quantity": "sum"})
        .reset_index()
    )
    product_sales["profit_margin"] = (
        product_sales["profit"] / product_sales["revenue"] * 100
    ).round(2)
    product_sales = product_sales.sort_values("revenue", ascending=False)

    print("\nSales by product category:")
    print(product_sales)

    # 3.2: Sales by region and customer segment
    region_segment_sales = (
        data.groupby(["region", "customer_segment"])
        .agg({"revenue": "sum"})
        .reset_index()
    )
    region_segment_pivot = region_segment_sales.pivot(
        index="region", columns="customer_segment", values="revenue"
    ).reset_index()

    print("\nSales by region and customer segment:")
    print(region_segment_pivot)

    # 3.3: Monthly sales trend
    monthly_sales = (
        data.groupby(["year", "month", "month_name"])
        .agg({"revenue": "sum", "profit": "sum"})
        .reset_index()
    )

    # Create a proper sort order for months
    month_order = {month: i for i, month in enumerate(calendar.month_abbr[1:])}
    monthly_sales["month_order"] = monthly_sales["month_name"].map(month_order)
    monthly_sales = monthly_sales.sort_values(["year", "month_order"])
    monthly_sales["YM"] = (
        monthly_sales["month_name"] + " " + monthly_sales["year"].astype(str)
    )

    print("\nMonthly sales trend:")
    print(monthly_sales.head(12))

    # 3.4: Quarterly sales with year-over-year comparison
    quarterly_sales = (
        data.groupby(["year", "quarter_name"])
        .agg({"revenue": "sum", "profit": "sum"})
        .reset_index()
    )

    # Pivot for year-over-year comparison
    quarterly_pivot = quarterly_sales.pivot(
        index="quarter_name", columns="year", values="revenue"
    ).reset_index()

    # Calculate year-over-year growth if we have multiple years
    if len(quarterly_sales["year"].unique()) > 1:
        years = sorted(quarterly_sales["year"].unique())
        for i in range(1, len(years)):
            prev_year = years[i - 1]
            curr_year = years[i]
            quarterly_pivot[f"YoY_Growth_{curr_year}"] = (
                (quarterly_pivot[curr_year] / quarterly_pivot[prev_year] - 1) * 100
            ).round(2)

    print("\nQuarterly sales with year-over-year comparison:")
    print(quarterly_pivot)

    # 3.5: Sales by channel with performance metrics
    channel_metrics = (
        data.groupby("sales_channel")
        .agg(
            {
                "transaction_id": "nunique",  # Count of transactions
                "customer_id": "nunique",  # Count of unique customers
                "revenue": "sum",
                "profit": "sum",
                "quantity": "sum",
            }
        )
        .reset_index()
    )

    # Calculate derived metrics
    channel_metrics["avg_order_value"] = (
        channel_metrics["revenue"] / channel_metrics["transaction_id"]
    ).round(2)
    channel_metrics["profit_margin"] = (
        channel_metrics["profit"] / channel_metrics["revenue"] * 100
    ).round(2)
    channel_metrics["revenue_per_customer"] = (
        channel_metrics["revenue"] / channel_metrics["customer_id"]
    ).round(2)

    print("\nSales channel metrics:")
    print(channel_metrics)

    # 3.6: Top selling products
    product_sales_detail = (
        data.groupby(["product_category", "product_name"])
        .agg({"quantity": "sum", "revenue": "sum", "profit": "sum"})
        .reset_index()
    )
    product_sales_detail["profit_margin"] = (
        product_sales_detail["profit"] / product_sales_detail["revenue"] * 100
    ).round(2)
    top_products = product_sales_detail.sort_values("revenue", ascending=False).head(10)

    print("\nTop 10 products by revenue:")
    print(top_products)

    # 3.7: Daily sales pattern by day of week
    day_of_week_sales = (
        data.groupby("day_of_week")
        .agg({"revenue": "sum", "transaction_id": pd.Series.nunique})
        .reset_index()
    )

    # Set correct order for days of week
    day_order = {
        "Monday": 0,
        "Tuesday": 1,
        "Wednesday": 2,
        "Thursday": 3,
        "Friday": 4,
        "Saturday": 5,
        "Sunday": 6,
    }
    day_of_week_sales["day_order"] = day_of_week_sales["day_of_week"].map(day_order)
    day_of_week_sales = day_of_week_sales.sort_values("day_order")
    day_of_week_sales["avg_transaction"] = (
        day_of_week_sales["revenue"] / day_of_week_sales["transaction_id"]
    ).round(2)

    print("\nSales by day of week:")
    print(
        day_of_week_sales[
            ["day_of_week", "revenue", "transaction_id", "avg_transaction"]
        ]
    )

    # -----------------------------------------
    # Step 4: Prepare data for specific chart types
    # -----------------------------------------

    # 4.1: Data for donut chart (product category breakdown)
    donut_data = product_sales[["product_category", "revenue"]].copy()
    # Sort by value for better visualization
    donut_data = donut_data.sort_values("revenue", ascending=False)

    # 4.2: Data for stacked bar chart (region and segment)
    stacked_data = region_segment_pivot.copy()

    # 4.3: Data for line chart (monthly trend)
    line_data = (
        monthly_sales[["YM", "revenue", "profit"]].tail(12).copy()
    )  # Last 12 months

    # 4.4: Data for clustered bar chart (quarterly comparison)
    if (
        "YoY_Growth_2023" in quarterly_pivot.columns
    ):  # Assuming 2023 is our current year
        bar_data = quarterly_pivot[
            ["quarter_name", 2022, 2023, "YoY_Growth_2023"]
        ].copy()
    else:
        # If we don't have YoY data, just use the latest year
        latest_year = quarterly_sales["year"].max()
        bar_data = quarterly_sales[quarterly_sales["year"] == latest_year][
            ["quarter_name", "revenue"]
        ]

    # 4.5: Data for a complex table with conditional formatting
    table_data = channel_metrics.copy()
    table_data = table_data.sort_values("revenue", ascending=False)

    # -----------------------------------------
    # Step 5: Create PowerPoint presentation
    # -----------------------------------------
    # Create the presentation using the prepared data
    create_sales_presentation(
        product_sales=product_sales,
        region_segment_pivot=region_segment_pivot,
        monthly_sales=line_data,
        quarterly_data=quarterly_pivot,
        channel_metrics=channel_metrics,
        top_products=top_products,
        day_of_week_sales=day_of_week_sales,
    )

    return data, product_sales  # Return processed data for further analysis if needed


def generate_sample_sales_data(num_records=1000):
    """Generate a sample sales dataset for demonstration."""
    np.random.seed(42)  # For reproducibility

    # Date range for the last 2 years
    end_date = datetime.now()
    start_date = end_date - timedelta(days=730)  # Approximately 2 years

    # Generate random dates
    dates = [
        start_date + timedelta(days=np.random.randint(0, 730))
        for _ in range(num_records)
    ]

    # Generate random transaction IDs
    transaction_ids = [f"T{i+10000}" for i in range(num_records)]

    # Generate random customer IDs (fewer than transactions to simulate repeat customers)
    customer_ids = [f"C{np.random.randint(1000, 2000)}" for _ in range(num_records)]

    # Product categories and names
    product_categories = [
        "Electronics",
        "Clothing",
        "Home & Kitchen",
        "Books",
        "Sports",
    ]
    product_names = {
        "Electronics": ["Smartphone", "Laptop", "Tablet", "Headphones", "Smartwatch"],
        "Clothing": ["T-Shirt", "Jeans", "Dress", "Jacket", "Shoes"],
        "Home & Kitchen": [
            "Blender",
            "Coffee Maker",
            "Toaster",
            "Cookware Set",
            "Cutlery",
        ],
        "Books": ["Fiction", "Non-Fiction", "Biography", "Cookbook", "Self-Help"],
        "Sports": [
            "Running Shoes",
            "Yoga Mat",
            "Fitness Tracker",
            "Weights",
            "Water Bottle",
        ],
    }

    # Generate product category and name pairs
    categories = [np.random.choice(product_categories) for _ in range(num_records)]
    products = [np.random.choice(product_names[category]) for category in categories]

    # Unit prices and costs by product category
    price_ranges = {
        "Electronics": (100, 1200),
        "Clothing": (15, 150),
        "Home & Kitchen": (25, 300),
        "Books": (10, 50),
        "Sports": (20, 200),
    }

    # Generate unit prices based on category
    unit_prices = [
        np.random.uniform(price_ranges[cat][0], price_ranges[cat][1])
        for cat in categories
    ]

    # Generate unit costs (60-80% of price)
    unit_costs = [price * np.random.uniform(0.6, 0.8) for price in unit_prices]

    # Generate quantities (1-5 items per transaction)
    quantities = [np.random.randint(1, 6) for _ in range(num_records)]

    # Regions
    regions = ["North", "South", "East", "West", "Central"]

    # Sales channels
    channels = ["Online", "Retail Store", "Phone", "Distributor"]

    # Customer segments
    segments = [
        "Consumer",
        "Business",
        "Government",
        np.nan,
    ]  # Include some missing values
    segment_weights = [0.7, 0.2, 0.05, 0.05]  # Probabilities for each segment

    # Create the DataFrame
    data = pd.DataFrame(
        {
            "transaction_id": transaction_ids,
            "transaction_date": dates,
            "customer_id": customer_ids,
            "customer_segment": np.random.choice(
                segments, size=num_records, p=segment_weights
            ),
            "region": [np.random.choice(regions) for _ in range(num_records)],
            "sales_channel": [np.random.choice(channels) for _ in range(num_records)],
            "product_category": categories,
            "product_name": products,
            "quantity": quantities,
            "unit_price": [round(price, 2) for price in unit_prices],
            "unit_cost": [round(cost, 2) for cost in unit_costs],
        }
    )

    return data


def create_sales_presentation(
    product_sales,
    region_segment_pivot,
    monthly_sales,
    quarterly_data,
    channel_metrics,
    top_products,
    day_of_week_sales,
):
    """Create a PowerPoint presentation with the processed sales data."""
    prs = PPTXPresentation()

    # Create title slide
    title_slide = prs.add_slide(layout_type=0, title="Sales Performance Analysis")
    title_slide.add_text_box(
        "Data-Driven Insights",
        position=(1, 3.5),
        size=(8, 0.8),
        align="center",
        font_size=20,
        bold=True,
    )

    # Create product category slide
    product_slide = prs.add_slide(title="Sales by Product Category")

    # Add a bar chart for product category sales
    product_slide.add_chart(
        chart_type="bar",
        data=product_sales,
        position=(0.5, 1.2),
        size=(5, 3.5),
        title="Revenue by Product Category",
        category_column="product_category",
        series_columns=["revenue"],
        series_1_name="Revenue",
        series_1_color="4472C4",  # Blue
        has_data_labels=True,
        data_label_number_format="$#,##0,K",
    )

    # Add a donut chart for product category contribution
    product_slide.add_chart(
        chart_type="donut",
        data=product_sales,
        position=(6, 1.2),
        size=(4, 3.5),
        title="Revenue Distribution",
        category_column="product_category",
        value_column="revenue",
        series_name="Revenue",
        has_legend=True,
        legend_position="right",
        data_label_number_format="0%",
    )

    # Add a table with product category data
    product_slide.add_table(
        data=product_sales,
        position=(0.5, 5),
        table_width=9.5,
        column_formats={
            "product_category": "text",
            "revenue": "dollars",
            "profit": "dollars",
            "quantity": "counts",
            "profit_margin": "percentage",
        },
        has_header=True,
        header_style={"fill_color": "4472C4", "font_color": "FFFFFF"},
    )

    # Apply conditional formatting to the table
    product_slide.apply_conditional_formatting(
        product_slide.tables[-1],
        rules=[
            {
                "type": "color_scale",
                "column": "profit_margin",
                "min_color": "F8696B",  # Red
                "max_color": "63BE7B",  # Green
            }
        ],
        start_row=1,
    )

    # Create regional sales slide
    # First, ensure the pivot table is in a format suitable for charting
    # We might need to reset the index and possibly melt the data
    region_data = region_segment_pivot.copy()

    # If we have a proper pivot table with columns for each segment
    if "region" in region_data.columns and "Consumer" in region_data.columns:
        # We have a proper pivot table, we can use it directly
        region_slide = prs.add_slide(title="Sales by Region and Customer Segment")

        # Add a stacked bar chart
        region_slide.add_chart(
            chart_type="stacked_bar",
            data=region_data,
            position=(0.5, 1.2),
            size=(9, 4),
            title="Revenue by Region and Customer Segment",
            category_column="region",
            series_columns=[col for col in region_data.columns if col != "region"],
            series_1_color="4472C4",  # Blue
            series_2_color="ED7D31",  # Orange
            series_3_color="A5A5A5",  # Gray
            data_label_position="inside_end",
            data_label_font_color="FFFFFF",
            data_label_number_format="$#,##0,K",
        )

    # Create monthly trend slide
    trend_slide = prs.add_slide(title="Monthly Sales Trend")

    # Add a line chart for monthly trend
    trend_slide.add_chart(
        chart_type="line",
        data=monthly_sales,
        position=(0.5, 1.2),
        size=(9, 4),
        title="Revenue and Profit Trend",
        category_column="YM",
        series_columns=["revenue", "profit"],
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

    # Create sales channel analysis slide
    channel_slide = prs.add_slide(title="Sales Channel Performance")

    # Add a clustered bar chart for channel metrics
    channel_slide.add_chart(
        chart_type="clustered_bar",
        data=channel_metrics,
        position=(0.5, 1.2),
        size=(9, 3),
        title="Revenue and Profit by Sales Channel",
        category_column="sales_channel",
        series_columns=["revenue", "profit"],
        series_1_name="Revenue",
        series_2_name="Profit",
        series_1_color="4472C4",  # Blue
        series_2_color="70AD47",  # Green
        has_data_labels=True,
        data_label_number_format="$#,##0,K",
    )

    # Add a table with channel metrics
    channel_display = channel_metrics[
        [
            "sales_channel",
            "transaction_id",
            "customer_id",
            "avg_order_value",
            "revenue_per_customer",
            "profit_margin",
        ]
    ].copy()

    # Rename columns for better presentation
    channel_display.columns = [
        "Sales Channel",
        "Transactions",
        "Customers",
        "Avg Order Value",
        "Revenue per Customer",
        "Profit Margin",
    ]

    channel_slide.add_table(
        data=channel_display,
        position=(0.5, 4.5),
        table_width=9.5,
        column_formats={
            "Sales Channel": "text",
            "Transactions": "counts",
            "Customers": "counts",
            "Avg Order Value": "dollars",
            "Revenue per Customer": "dollars",
            "Profit Margin": "percentage",
        },
        has_header=True,
        header_style={"fill_color": "4472C4", "font_color": "FFFFFF"},
    )

    # Apply conditional formatting for the channel metrics
    channel_slide.apply_conditional_formatting(
        channel_slide.tables[-1],
        rules=[
            {
                "type": "color_scale",
                "column": "Profit Margin",
                "min_color": "FFFFFF",  # White
                "max_color": "63BE7B",  # Green
            },
            {
                "type": "color_scale",
                "column": "Avg Order Value",
                "min_color": "FFFFFF",  # White
                "max_color": "4472C4",  # Blue
            },
        ],
        start_row=1,
    )

    # Create day of week analysis slide
    daily_slide = prs.add_slide(title="Sales Pattern by Day of Week")

    # Add a bar chart for day of week sales
    daily_slide.add_chart(
        chart_type="bar",
        data=day_of_week_sales,
        position=(0.5, 1.2),
        size=(4.5, 3),
        title="Revenue by Day of Week",
        category_column="day_of_week",
        series_columns=["revenue"],
        series_1_name="Revenue",
        series_1_color="4472C4",  # Blue
        has_data_labels=True,
        data_label_number_format="$#,##0,K",
    )

    # Add a line chart for average transaction value
    daily_slide.add_chart(
        chart_type="line",
        data=day_of_week_sales,
        position=(5.5, 1.2),
        size=(4.5, 3),
        title="Average Transaction Value",
        category_column="day_of_week",
        series_columns=["avg_transaction"],
        series_1_name="Avg Transaction",
        series_1_color="ED7D31",  # Orange
        series_1_line_width=2.5,
        show_markers=True,
        has_data_labels=True,
        data_label_number_format="$#,##0",
    )

    # Add a table with day of week data
    daily_display = day_of_week_sales[
        ["day_of_week", "transaction_id", "revenue", "avg_transaction"]
    ].copy()
    daily_display.columns = [
        "Day of Week",
        "Transactions",
        "Revenue",
        "Avg Transaction",
    ]

    daily_slide.add_table(
        data=daily_display,
        position=(0.5, 4.5),
        table_width=9.5,
        column_formats={
            "Day of Week": "text",
            "Transactions": "counts",
            "Revenue": "dollars",
            "Avg Transaction": "dollars",
        },
        has_header=True,
        header_style={"fill_color": "4472C4", "font_color": "FFFFFF"},
    )

    # Create top products slide
    top_slide = prs.add_slide(title="Top Performing Products")

    # Add a horizontal bar chart for top products
    # First, prepare the data to show only product name and category
    top_display = (
        top_products[["product_name", "product_category", "revenue"]].head(8).copy()
    )
    top_display["product"] = (
        top_display["product_name"] + " (" + top_display["product_category"] + ")"
    )
    top_chart_data = top_display[["product", "revenue"]].sort_values("revenue")

    top_slide.add_chart(
        chart_type="clustered_bar",
        data=top_chart_data,
        position=(0.5, 1.2),
        size=(9, 3),
        title="Top Products by Revenue",
        category_column="product",
        series_columns=["revenue"],
        series_1_name="Revenue",
        series_1_color="4472C4",  # Blue
        chart_type="bar_clustered",  # Horizontal bars
        has_data_labels=True,
        data_label_number_format="$#,##0,K",
    )

    # Add a table with top products data
    top_table_data = top_products[
        [
            "product_name",
            "product_category",
            "quantity",
            "revenue",
            "profit",
            "profit_margin",
        ]
    ].copy()
    top_table_data.columns = [
        "Product",
        "Category",
        "Units Sold",
        "Revenue",
        "Profit",
        "Profit Margin",
    ]

    top_slide.add_table(
        data=top_table_data,
        position=(0.5, 4.5),
        table_width=9.5,
        column_formats={
            "Product": "text",
            "Category": "text",
            "Units Sold": "counts",
            "Revenue": "dollars",
            "Profit": "dollars",
            "Profit Margin": "percentage",
        },
        has_header=True,
        header_style={"fill_color": "4472C4", "font_color": "FFFFFF"},
    )

    # Apply conditional formatting to top products table
    top_slide.apply_conditional_formatting(
        top_slide.tables[-1],
        rules=[
            {
                "type": "color_scale",
                "column": "Profit Margin",
                "min_color": "F8696B",  # Red
                "max_color": "63BE7B",  # Green
            },
            {
                "type": "top_bottom",
                "column": "Revenue",
                "top": True,
                "rank": 3,
                "color": "D8E4BC",  # Light green
            },
        ],
        start_row=1,
    )

    # Save the presentation
    prs.save("sales_performance_analysis.pptx")
    print("Sales performance presentation created: sales_performance_analysis.pptx")


if __name__ == "__main__":
    # Run the data preparation and visualization
    processed_data, product_data = prepare_and_visualize_sales_data()
    print("\nData preparation and PowerPoint creation complete!")
