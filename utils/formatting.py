"""
Number formatting utilities for the pptx_charts_tables package.
"""


def format_value(value, format_type, config):
    """
    Format a value based on format type.

    Args:
        value: The value to format.
        format_type (str): Type of formatting to apply ('dollars', 'percentage', 'counts', 'text').
        config (dict): Configuration settings.

    Returns:
        str: Formatted value as string.
    """
    if value is None or (isinstance(value, (float, int)) and pd.isna(value)):
        return ""

    if format_type == "text" or isinstance(value, str):
        return str(value)

    if format_type == "dollars":
        return format_dollars(value, config)
    elif format_type == "percentage":
        return format_percentage(value, config)
    elif format_type == "counts":
        return format_counts(value, config)
    else:
        return str(value)


def format_dollars(value, config):
    """
    Format a value as dollars.

    Args:
        value: The value to format.
        config (dict): Configuration settings.

    Returns:
        str: Formatted value as string.
    """
    if not isinstance(value, (int, float)) or pd.isna(value):
        return ""

    dollars_config = config["formatting"]["dollars"]

    # Apply scaling if configured
    scaling = dollars_config.get("scaling")
    if scaling == "K" and abs(value) >= 1000:
        value = value / 1000
        suffix = "K"
    elif scaling == "M" and abs(value) >= 1000000:
        value = value / 1000000
        suffix = "M"
    elif scaling == "B" and abs(value) >= 1000000000:
        value = value / 1000000000
        suffix = "B"
    else:
        suffix = ""

    # Format the number with configured decimal places
    decimal_places = dollars_config.get("decimal_places", 0)

    # Handle negative values with parentheses if configured
    negative_in_parentheses = dollars_config.get("negative_in_parentheses", True)

    if value < 0 and negative_in_parentheses:
        formatted = f"({format_number(abs(value), decimal_places)})"
    else:
        formatted = format_number(value, decimal_places)

    # Add dollar symbol if configured
    show_symbol = dollars_config.get("show_symbol", True)
    if show_symbol:
        return f"${formatted}{suffix}"
    else:
        return f"{formatted}{suffix}"


def format_percentage(value, config):
    """
    Format a value as percentage.

    Args:
        value: The value to format.
        config (dict): Configuration settings.

    Returns:
        str: Formatted value as string.
    """
    if not isinstance(value, (int, float)) or pd.isna(value):
        return ""

    percentage_config = config["formatting"]["percentages"]

    # Convert to percentage if needed
    if abs(value) <= 1:
        value = value * 100

    # Format the number with configured decimal places
    decimal_places = percentage_config.get("decimal_places", 1)

    # Handle negative values with parentheses if configured
    negative_in_parentheses = percentage_config.get("negative_in_parentheses", True)

    if value < 0 and negative_in_parentheses:
        formatted = f"({format_number(abs(value), decimal_places)})"
    else:
        formatted = format_number(value, decimal_places)

    # Add percentage symbol if configured
    show_symbol = percentage_config.get("show_symbol", True)
    if show_symbol:
        return f"{formatted}%"
    else:
        return formatted


def format_counts(value, config):
    """
    Format a value as a count (integer with commas).

    Args:
        value: The value to format.
        config (dict): Configuration settings.

    Returns:
        str: Formatted value as string.
    """
    if not isinstance(value, (int, float)) or pd.isna(value):
        return ""

    counts_config = config["formatting"]["counts"]

    # Round to integer
    value = round(value)

    # Format the number with commas if configured
    show_commas = counts_config.get("show_commas", True)

    # Handle negative values with parentheses if configured
    negative_in_parentheses = counts_config.get("negative_in_parentheses", True)

    if value < 0 and negative_in_parentheses:
        if show_commas:
            formatted = f"({format_number(abs(value), 0, show_commas=True)})"
        else:
            formatted = f"({int(abs(value))})"
    else:
        if show_commas:
            formatted = format_number(value, 0, show_commas=True)
        else:
            formatted = str(int(value))

    return formatted


def format_number(value, decimal_places, show_commas=True):
    """
    Format a number with specified decimal places and comma separators.

    Args:
        value: The value to format.
        decimal_places (int): Number of decimal places.
        show_commas (bool): Whether to use comma separators.

    Returns:
        str: Formatted number.
    """
    if show_commas:
        if decimal_places == 0:
            return f"{int(value):,}"
        else:
            # Format with decimal places and commas
            return f"{value:,.{decimal_places}f}"
    else:
        if decimal_places == 0:
            return str(int(value))
        else:
            return f"{value:.{decimal_places}f}"


# Add import for pandas at the top
import pandas as pd
