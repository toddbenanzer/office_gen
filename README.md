# pptx_charts_tables

A Python package for creating PowerPoint presentations with embedded charts and tables from pandas DataFrames, specially designed for financial and business reporting.

## Features

- Create PowerPoint presentations with embedded charts (native PowerPoint charts, not images)
- Supported chart types: Bar, Clustered Bar, Stacked Bar, Donut, and Line charts
- Styled tables with custom formatting for financial data
- Number formatting for dollars, percentages, and counts
- Fully customizable colors, fonts, borders, and other styling options
- Conditional formatting for tables (color scales, data bars, highlighting, top/bottom rules)
- Text boxes, shapes, connectors, and images with styling options
- Color utilities for creating harmonious color schemes and palettes
- Configure defaults with a configuration file
- Full control over positioning of elements on slides

## Installation

```bash
pip install pptx_charts_tables
```

## Dependencies

- pandas
- python-pptx

## Quick Start

```python
import pandas as pd
from pptx_charts_tables import PPTXPresentation

# Create sample data
quarterly_data = pd.DataFrame({
    'quarter': ['Q1', 'Q2', 'Q3', 'Q4'],
    'revenue': [1250000, 1450000, 1550000, 1750000],
    'costs': [950000, 1050000, 1150000, 1250000],
    'profit': [300000, 400000, 400000, 500000]
})

# Create presentation
prs = PPTXPresentation()

# Add a slide
slide = prs.add_slide(title="Quarterly Financial Performance")

# Add a chart
slide.add_chart(
    chart_type='clustered_bar',
    data=quarterly_data,
    position=(1, 1.5),
    size=(8, 4.5),
    title="Quarterly Revenue, Costs, and Profit",
    category_column='quarter',
    series_columns=['revenue', 'costs', 'profit'],
    series_1_name="Revenue",
    series_2_name="Costs",
    series_3_name="Profit",
    series_1_color="4472C4",  # Blue
    series_2_color="ED7D31",  # Orange
    series_3_color="70AD47",  # Green
    has_data_labels=True
)

# Add a table
table = slide.add_table(
    data=quarterly_data,
    position=(1, 6.25),
    table_width=8,
    column_formats={
        'quarter': 'text',
        'revenue': 'dollars',
        'costs': 'dollars',
        'profit': 'dollars'
    },
    has_header=True
)

# Apply conditional formatting to the table
slide.apply_conditional_formatting(
    table,
    rules=[
        {
            'type': 'color_scale',
            'column': 'profit',
            'min_color': 'F8696B',  # Red
            'max_color': '63BE7B'   # Green
        }
    ],
    start_row=1  # Skip header row
)

# Add a text box for additional information
slide.add_text_box(
    "This quarter's performance exceeded expectations with significant growth in all key metrics.",
    position=(1, 8),
    size=(8, 0.75),
    font_size=12,
    align='center',
    font_name='Calibri'
)

# Save the presentation
prs.save("financial_report.pptx")
```

## Supported Chart Types

- `bar` - Simple bar chart (vertical columns)
- `clustered_bar` - Bar chart with multiple series grouped
- `stacked_bar` - Bar chart with stacked series
- `donut` - Donut chart (pie chart with a hole)
- `line` - Line chart, with or without markers

## Number Formatting

The package supports the following number formatting options:

- `dollars` - Currency values with dollar sign, thousands separators, optional decimal places
- `percentage` - Percentage values with percent sign, optional decimal places
- `counts` - Integer values with thousands separators

## Customization

All visual elements are fully customizable:

- Colors for chart series, fills, borders, text
- Fonts, font sizes, font styles
- Data label positions and formats
- Table styling including headers, alternating rows, totals
- Chart types and subtypes
- Axis visibility and gridlines
- And much more

## Configuration

You can provide a custom configuration dictionary to override the default settings:

```python
custom_config = {
    'general': {
        'font_name': 'Calibri',
        'font_size': 10,
    },
    'charts': {
        'colors': {
            'series_1': 'FF0000',  # Red
            'series_2': '00FF00',  # Green
            'series_3': '0000FF',  # Blue
        }
    }
}

prs = PPTXPresentation(config=custom_config)
```

## Conditional Formatting

Apply conditional formatting to your tables to highlight important data insights:

```python
# Apply conditional formatting to a table
slide.apply_conditional_formatting(
    table,
    rules=[
        # Color scale (gradient)
        {
            'type': 'color_scale',
            'column': 'profit',
            'min_color': 'F8696B',  # Red
            'max_color': '63BE7B'   # Green
        },
        # Highlight cells
        {
            'type': 'highlight_cells',
            'column': 'growth',
            'operator': 'greater_than',
            'value': 15,
            'color': 'D8E4BC'  # Light green
        },
        # Top/bottom highlighting
        {
            'type': 'top_bottom',
            'column': 'revenue',
            'top': True,  # Top values (use False for bottom)
            'rank': 3,    # Number of values to highlight
            'color': 'B4C6E7'  # Light blue
        }
    ],
    start_row=1  # Skip header row
)
```

## Shapes and Text Boxes

Add shapes, text boxes, arrows, and images to your slides:

```python
# Add a text box
slide.add_text_box(
    "Important insight about this data",
    position=(1, 5),
    size=(4, 1),
    font_size=14,
    bold=True,
    align='center',
    fill_color='DDEBF7',
    border_color='5B9BD5'
)

# Add a shape
slide.add_shape(
    'rectangle',
    position=(6, 5),
    size=(2, 1),
    fill_color='70AD47',
    text="↑ 15%",
    font_size=18,
    font_color='FFFFFF',
    align='center',
    v_align='middle'
)

# Add an arrow
slide.add_arrow(
    start_pos=(5, 3),
    end_pos=(6, 3.5),
    color='ED7D31',
    width=2,
    dash_style='dash'
)

# Add an image
slide.add_image(
    'logo.png',
    position=(9, 0.5),
    size=(1, 0.5)
)
```

## Color Utilities

Generate harmonious color schemes:

```python
# Create a color palette based on a corporate color
colors = slide.create_color_palette(
    base_color='4472C4',  # Corporate blue
    variations=5,
    mode='monochromatic'
)

# Get a predefined color scheme
financial_colors = slide.get_color_scheme('financial')
```

## Advanced Examples

The package includes two example scripts:

1. `example_usage.py` - Basic examples of charts and tables
2. `advanced_financial_example.py` - Comprehensive financial report with multiple slides

The advanced example demonstrates:

- Creating a multi-slide financial report
- Using various chart types for different data visualizations
- Applying conditional formatting to tables
- Adding shapes, text boxes, and other visual elements
- Creating professional-looking financial metrics displays
- Working with various data types and formatting options

## Use Cases

This package is especially well-suited for:

- Financial reporting and dashboards
- Sales and marketing performance reports
- Business reviews and executive summaries
- Investor presentations
- Product performance analysis
- Budget presentations
- Quarterly business reviews

## License

MIT