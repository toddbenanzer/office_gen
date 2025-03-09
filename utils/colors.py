"""
Color utilities for the pptx_charts_tables package.
Provides functions for working with colors, color schemes, and gradients.
"""


def rgb_to_hex(r, g, b):
    """
    Convert RGB values to hex color string.
    
    Args:
        r (int): Red value (0-255)
        g (int): Green value (0-255)
        b (int): Blue value (0-255)
        
    Returns:
        str: Hex color string (e.g. 'FF5733')
    """
    return f"{r:02X}{g:02X}{b:02X}"


def hex_to_rgb(hex_str):
    """
    Convert hex color string to RGB tuple.
    
    Args:
        hex_str (str): Hex color string (e.g. 'FF5733')
        
    Returns:
        tuple: (r, g, b) tuple with values 0-255
    """
    # Remove # if present
    hex_str = hex_str.lstrip('#')
    
    # Convert to RGB
    return tuple(int(hex_str[i:i+2], 16) for i in (0, 2, 4))


def create_color_scale(start_color, end_color, steps):
    """
    Create a color scale between two colors.
    
    Args:
        start_color (str): Starting hex color
        end_color (str): Ending hex color
        steps (int): Number of colors in the scale
        
    Returns:
        list: List of hex color strings
    """
    start_rgb = hex_to_rgb(start_color)
    end_rgb = hex_to_rgb(end_color)
    
    result = []
    for i in range(steps):
        # Calculate the proportion of the step
        proportion = i / (steps - 1) if steps > 1 else 0
        
        # Interpolate between the colors
        r = round(start_rgb[0] + proportion * (end_rgb[0] - start_rgb[0]))
        g = round(start_rgb[1] + proportion * (end_rgb[1] - start_rgb[1]))
        b = round(start_rgb[2] + proportion * (end_rgb[2] - start_rgb[2]))
        
        result.append(rgb_to_hex(r, g, b))
    
    return result


def create_palette(base_color, variations=5, mode='monochromatic'):
    """
    Create a color palette based on a base color.
    
    Args:
        base_color (str): Base hex color
        variations (int): Number of variations
        mode (str): Type of palette - 'monochromatic', 'complementary', 'analogous'
        
    Returns:
        list: List of hex color strings
    """
    # Convert base color to RGB
    base_rgb = hex_to_rgb(base_color)
    
    if mode == 'monochromatic':
        # Create monochromatic variations (shades and tints)
        return create_monochromatic_palette(base_color, variations)
    
    elif mode == 'complementary':
        # Create a palette with base color and its complement
        return create_complementary_palette(base_color, variations)
    
    elif mode == 'analogous':
        # Create analogous colors
        return create_analogous_palette(base_color, variations)
    
    else:
        # Default to monochromatic
        return create_monochromatic_palette(base_color, variations)


def create_monochromatic_palette(base_color, variations=5):
    """
    Create a monochromatic palette (lighter and darker shades).

    Args:
        base_color (str): Base hex color
        variations (int): Number of variations

    Returns:
        list: List of hex color strings
    """
    # Start with white and end with the base color for lighter shades
    lighter_shades = create_color_scale('FFFFFF', base_color, variations // 2 + 1)

    # Start with the base color and end with black for darker shades
    darker_shades = create_color_scale(base_color, '000000', variations // 2 + 1)

    # Combine, but avoid duplicating the base color
    return lighter_shades[:-1] + darker_shades


def create_complementary_palette(base_color, variations=5):
    """
    Create a palette with a base color and its complement.

    Args:
        base_color (str): Base hex color
        variations (int): Number of variations

    Returns:
        list: List of hex color strings
    """
    # Convert to RGB
    r, g, b = hex_to_rgb(base_color)

    # Calculate complement (invert each component)
    complement = rgb_to_hex(255 - r, 255 - g, 255 - b)

    # Create scales between the colors
    return create_color_scale(base_color, complement, variations)


def create_analogous_palette(base_color, variations=5):
    """
    Create an analogous color palette.
    This is a simplified version as true analogous colors require HSL/HSV conversion.

    Args:
        base_color (str): Base hex color
        variations (int): Number of variations

    Returns:
        list: List of hex color strings
    """
    # Convert to RGB
    r, g, b = hex_to_rgb(base_color)

    # Create variations by rotating RGB values
    result = [base_color]

    for i in range(1, variations):
        # Rotate RGB values (simplified approach)
        r_new = (r + i * 30) % 256
        g_new = (g + i * 20) % 256
        b_new = (b + i * 10) % 256

        result.append(rgb_to_hex(r_new, g_new, b_new))

    return result


def get_common_color_schemes():
    """
    Return a dictionary of common color schemes.

    Returns:
        dict: Dictionary of color scheme name to list of hex colors
    """
    return {
        'blue': ['4472C4', '5B9BD5', '8FAADC', 'B4C7E7', 'D9E1F2'],
        'green': ['70AD47', '9BBB59', 'A9D08E', 'C5E0B4', 'E2EFD9'],
        'red': ['C00000', 'FF0000', 'FF6666', 'FF9999', 'FFCCCC'],
        'orange': ['ED7D31', 'F4B183', 'F8CBAD', 'FCE4D6', 'FFF2CC'],
        'purple': ['7030A0', '8064A2', '9B82BB', 'B2A1C7', 'CCC0DA'],
        'grayscale': ['000000', '444444', '888888', 'BBBBBB', 'EEEEEE'],
        'pastel': ['FFCCCC', 'FFEBCC', 'FFFFCC', 'EBFFCC', 'CCFFCC', 
                   'CCFFEB', 'CCFFFF', 'CCEBFF', 'CCCCFF', 'EBCCFF'],
        'contrast': ['004489', 'E8BD00', 'A40122', '53A2BE', '15846B',
                     'AA57AA', 'F5793A', '0BA02C', '333333', '8C8C8C'],
        'financial': ['3366CC', 'DC3912', 'FF9900', '109618', '990099',
                     '0099C6', 'DD4477', '66AA00', 'B82E2E', '316395']
    }


def get_color_scheme(scheme_name):
    """
    Get a specific color scheme by name.

    Args:
        scheme_name (str): Name of the color scheme

    Returns:
        list: List of hex color strings, or None if not found
    """
    schemes = get_common_color_schemes()
    return schemes.get(scheme_name.lower())