"""
Conditional formatting utilities for the pptx_charts_tables package.
Provides functions for applying conditional formatting to tables.
"""

from pptx.dml.color import RGBColor
from .colors import hex_to_rgb, create_color_scale


def apply_conditional_formatting(table, rules, start_row=0):
    """
    Apply conditional formatting to a table based on rules.

    Args:
        table: The table object
        rules (list): List of formatting rules
        start_row (int): Row to start applying formatting (to skip headers)
    """
    # Process each rule
    for rule in rules:
        rule_type = rule.get("type", "")

        if rule_type == "color_scale":
            apply_color_scale(table, rule, start_row)

        elif rule_type == "data_bar":
            apply_data_bars(table, rule, start_row)

        elif rule_type == "icon_set":
            # This is more complex in pptx, but we can approximate
            apply_icon_set(table, rule, start_row)

        elif rule_type == "highlight_cells":
            apply_highlight_cells(table, rule, start_row)

        elif rule_type == "top_bottom":
            apply_top_bottom(table, rule, start_row)


def apply_color_scale(table, rule, start_row):
    """
    Apply a color scale conditional format to a table.

    Args:
        table: The table object
        rule (dict): Color scale rule definition
        start_row (int): Row to start applying formatting
    """
    # Get rule parameters
    column = rule.get("column")
    min_color = rule.get("min_color", "63BE7B")  # Green
    mid_color = rule.get("mid_color")
    max_color = rule.get("max_color", "F8696B")  # Red

    # Find the column index
    col_idx = _find_column_index(table, column, rule)
    if col_idx is None:
        return

    # Get values from the column
    values = []
    for row_idx in range(start_row, len(table.rows)):
        cell = table.cell(row_idx, col_idx)
        # Try to convert to float, skip non-numeric cells
        try:
            value = float(cell.text.replace("$", "").replace(",", "").replace("%", ""))
            values.append((row_idx, value))
        except ValueError:
            continue

    if not values:
        return

    # Find min and max values
    min_val = min(values, key=lambda x: x[1])[1]
    max_val = max(values, key=lambda x: x[1])[1]

    # Create color scale
    if mid_color:
        # Create two scales: min to mid and mid to max
        mid_point = (min_val + max_val) / 2
        lower_scale = create_color_scale(min_color, mid_color, 50)
        upper_scale = create_color_scale(mid_color, max_color, 50)
        color_scale = lower_scale + upper_scale
        scale_min = min_val
        scale_max = max_val
        scale_range = scale_max - scale_min
    else:
        # Simple min to max scale
        color_scale = create_color_scale(min_color, max_color, 100)
        scale_min = min_val
        scale_max = max_val
        scale_range = scale_max - scale_min

    # Apply colors to cells
    for row_idx, value in values:
        if scale_range > 0:
            # Calculate position in scale (0-99)
            position = int(((value - scale_min) / scale_range) * 99)
            position = max(0, min(99, position))  # Ensure within bounds

            # Get color from scale
            color = color_scale[position]

            # Apply color to cell
            cell = table.cell(row_idx, col_idx)
            cell.fill.solid()
            cell.fill.fore_color.rgb = RGBColor.from_string(color)


def apply_data_bars(table, rule, start_row):
    """
    Apply data bars conditional format to a table.
    Since PowerPoint doesn't support data bars directly, we approximate with cell background.

    Args:
        table: The table object
        rule (dict): Data bar rule definition
        start_row (int): Row to start applying formatting
    """
    # Get rule parameters
    column = rule.get("column")
    color = rule.get("color", "638EC6")  # Blue

    # Find the column index
    col_idx = _find_column_index(table, column, rule)
    if col_idx is None:
        return

    # Get values from the column
    values = []
    for row_idx in range(start_row, len(table.rows)):
        cell = table.cell(row_idx, col_idx)
        # Try to convert to float, skip non-numeric cells
        try:
            value = float(cell.text.replace("$", "").replace(",", "").replace("%", ""))
            values.append((row_idx, value))
        except ValueError:
            continue

    if not values:
        return

    # Find min and max values
    min_val = min(values, key=lambda x: x[1])[1]
    max_val = max(values, key=lambda x: x[1])[1]

    # Ensure min_val is not equal to max_val to avoid division by zero
    if min_val == max_val:
        max_val = min_val + 1

    value_range = max_val - min_val

    # Apply data bars to cells
    for row_idx, value in values:
        cell = table.cell(row_idx, col_idx)

        # Calculate width percentage
        if value_range > 0:
            width_pct = (value - min_val) / value_range
        else:
            width_pct = 0.5  # Default to 50% if all values are the same

        # Create a gradient-like effect
        # Note: This is an approximation since PowerPoint doesn't support gradient fills
        # directly through the python-pptx API
        cell.fill.solid()
        cell.fill.fore_color.rgb = RGBColor.from_string(color)

        # We could potentially add rectangle shapes on top of cells for a better effect,
        # but that's more complex and beyond this implementation


def apply_icon_set(table, rule, start_row):
    """
    Approximate icon sets with formatting.
    Since PowerPoint doesn't support icon sets directly, we approximate with cell background.

    Args:
        table: The table object
        rule (dict): Icon set rule definition
        start_row (int): Row to start applying formatting
    """
    # Get rule parameters
    column = rule.get("column")
    thresholds = rule.get("thresholds", [33, 67])  # Default to thirds

    # Find the column index
    col_idx = _find_column_index(table, column, rule)
    if col_idx is None:
        return

    # Define colors for each threshold level
    colors = rule.get("colors", ["F8696B", "FFEB84", "63BE7B"])  # Red, Yellow, Green

    # Get values from the column
    values = []
    for row_idx in range(start_row, len(table.rows)):
        cell = table.cell(row_idx, col_idx)
        # Try to convert to float, skip non-numeric cells
        try:
            value = float(cell.text.replace("$", "").replace(",", "").replace("%", ""))
            values.append((row_idx, value))
        except ValueError:
            continue

    if not values:
        return

    # Find min and max values
    min_val = min(values, key=lambda x: x[1])[1]
    max_val = max(values, key=lambda x: x[1])[1]

    # Ensure min_val is not equal to max_val to avoid division by zero
    if min_val == max_val:
        max_val = min_val + 1

    value_range = max_val - min_val

    # Calculate actual threshold values
    threshold_values = [min_val + (value_range * t / 100) for t in thresholds]

    # Apply formatting based on thresholds
    for row_idx, value in values:
        cell = table.cell(row_idx, col_idx)

        # Determine which threshold the value falls under
        if value <= threshold_values[0]:
            color_idx = 0
        elif len(threshold_values) > 1 and value <= threshold_values[1]:
            color_idx = 1
        else:
            color_idx = 2

        # Apply the corresponding color
        cell.fill.solid()
        cell.fill.fore_color.rgb = RGBColor.from_string(colors[color_idx])


def apply_highlight_cells(table, rule, start_row):
    """
    Apply highlight cells conditional format to a table.

    Args:
        table: The table object
        rule (dict): Highlight cells rule definition
        start_row (int): Row to start applying formatting
    """
    # Get rule parameters
    column = rule.get("column")
    operator = rule.get("operator", "greater_than")
    value = rule.get("value", 0)
    color = rule.get("color", "FF0000")  # Red

    # Find the column index
    col_idx = _find_column_index(table, column, rule)
    if col_idx is None:
        return

    # Apply highlighting to cells
    for row_idx in range(start_row, len(table.rows)):
        cell = table.cell(row_idx, col_idx)

        # Try to convert to float, skip non-numeric cells
        try:
            cell_value = float(
                cell.text.replace("$", "").replace(",", "").replace("%", "")
            )
        except ValueError:
            continue

        # Check condition based on operator
        highlight = False

        if operator == "greater_than":
            highlight = cell_value > value
        elif operator == "less_than":
            highlight = cell_value < value
        elif operator == "equal_to":
            highlight = cell_value == value
        elif operator == "not_equal_to":
            highlight = cell_value != value
        elif operator == "greater_than_or_equal":
            highlight = cell_value >= value
        elif operator == "less_than_or_equal":
            highlight = cell_value <= value

        # Apply highlight if condition is met
        if highlight:
            cell.fill.solid()
            cell.fill.fore_color.rgb = RGBColor.from_string(color)


def apply_top_bottom(table, rule, start_row):
    """
    Apply top/bottom conditional format to a table.

    Args:
        table: The table object
        rule (dict): Top/bottom rule definition
        start_row (int): Row to start applying formatting
    """
    # Get rule parameters
    column = rule.get("column")
    top = rule.get("top", True)  # True for top, False for bottom
    percent = rule.get("percent", False)  # True for percent, False for count
    rank = rule.get("rank", 10)  # Top/bottom 10
    color = rule.get("color", "63BE7B")  # Green

    # Find the column index
    col_idx = _find_column_index(table, column, rule)
    if col_idx is None:
        return

    # Get values from the column
    values = []
    for row_idx in range(start_row, len(table.rows)):
        cell = table.cell(row_idx, col_idx)
        # Try to convert to float, skip non-numeric cells
        try:
            value = float(cell.text.replace("$", "").replace(",", "").replace("%", ""))
            values.append((row_idx, value))
        except ValueError:
            continue

    if not values:
        return

    # Sort values
    if top:
        values.sort(key=lambda x: x[1], reverse=True)  # Descending for top
    else:
        values.sort(key=lambda x: x[1])  # Ascending for bottom

    # Calculate how many items to highlight
    if percent:
        count = int(len(values) * rank / 100)
    else:
        count = min(rank, len(values))

    # Apply highlighting to top/bottom cells
    for i in range(count):
        if i < len(values):
            row_idx, _ = values[i]
            cell = table.cell(row_idx, col_idx)
            cell.fill.solid()
            cell.fill.fore_color.rgb = RGBColor.from_string(color)


def _find_column_index(table, column, rule):
    """
    Helper to find the column index.

    Args:
        table: The table object
        column: Column identifier (name or index)
        rule: Rule definition

    Returns:
        int: Column index or None if not found
    """
    # If column is an integer, use it directly
    if isinstance(column, int):
        if 0 <= column < len(table.columns):
            return column
        return None

    # If column is a string, try to find it in the header row
    col_idx = rule.get("col_idx")
    if col_idx is not None:
        if 0 <= col_idx < len(table.columns):
            return col_idx

    # Try to find the column by name in the header row
    if isinstance(column, str):
        for i in range(len(table.columns)):
            if table.cell(0, i).text == column:
                return i

    return None
