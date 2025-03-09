"""
Default configuration settings for the pptx_charts_tables package.
This module defines the default styles, colors, and formatting options.
"""

DEFAULT_CONFIG = {
    'general': {
        'font_name': 'Arial',
        'font_size': 11,
        'title_font_size': 14,
    },
    'charts': {
        'gap_width': 100,
        'has_data_labels': True,
        'data_label_font_name': 'Arial',
        'data_label_font_size': 11,
        'has_legend': True,
        'legend_position': 'bottom',  # 'bottom', 'right', 'left', 'top'
        'legend_font_name': 'Arial',
        'legend_font_size': 9,
        'axis_font_name': 'Arial',
        'axis_font_size': 9,
        'show_axis_lines': False,
        'show_gridlines': False,
        'value_axis_visible': False,
        'colors': {
            'series_1': '3C2F80',  # Purple
            'series_2': '2C1F10',  # Dark brown
            'series_3': '4C4C4C',  # Gray
            'series_4': '5C5F80',  # Blue-gray
            'series_5': '6C6F10',  # Olive green
        }
    },
    'bar_chart': {
        'chart_type': 'column_clustered',
        'gap_width': 100,
    },
    'clustered_bar_chart': {
        'chart_type': 'column_clustered',
        'gap_width': 100,
    },
    'stacked_bar_chart': {
        'chart_type': 'column_stacked',
        'gap_width': 100,
        'data_label_position': 'inside_end',
    },
    'donut_chart': {
        'data_label_number_format': '0%',
    },
    'line_chart': {
        'line_width': {
            'series_1': 2.5,
            'series_2': 1.5,
            'series_3': 1.5,
            'series_4': 1.5,
            'series_5': 1.5,
        },
        'line_style': {
            'series_1': 'solid',
            'series_2': 'dash',
            'series_3': 'solid',
            'series_4': 'dash',
            'series_5': 'solid',
        },
        'show_markers': False,
        'show_gridlines': True,
        'value_axis_visible': True,
    },
    'tables': {
        'has_header': True,
        'header_font_name': 'Arial',
        'header_font_size': 12,
        'header_font_bold': True,
        'header_fill_color': '3C2F80',
        'header_font_color': 'FFFFFF',
        'cell_font_name': 'Arial',
        'cell_font_size': 11,
        'row_height': 0.5,  # in inches
        'totals_font_bold': True,
        'totals_border_top': True,
        'alternating_row_fill': True,
        'alternating_row_fill_color': 'F2F2F2',
    },
    'formatting': {
        'dollars': {
            'decimal_places': 0,
            'show_symbol': True,
            'negative_in_parentheses': True,
            'negative_color': 'FF0000',
            'scaling': None,  # None, 'K', 'M', or 'B'
        },
        'percentages': {
            'decimal_places': 1,
            'show_symbol': True,
            'negative_in_parentheses': True,
            'negative_color': 'FF0000',
        },
        'counts': {
            'decimal_places': 0,
            'show_commas': True,
            'negative_in_parentheses': True,
            'negative_color': 'FF0000',
        }
    },
    'text_box': {
        'size': {
            'width': 7,
            'height': 1,
        },
        'font': {
            'name': 'Arial',
            'size': 11,
            'bold': False,
            'italic': False,
            'underline': False,
            'color': '000000',
        },
        'fill': {
            'color': 'FFFFFF',
            'transparency': 0,
            'no_fill': True,
        },
        'border': {
            'color': '000000',
            'weight': 1,
            'no_border': True,
        },
        'alignment': {
            'horizontal': 'center',
            'vertical': 'middle',
        },
        'margin': {
            'top': 0.1,
            'bottom': 0.1,
            'left': 0.1,
            'right': 0.1,
        },
    },
}