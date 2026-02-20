"""
Format dictionaries for invoice cell styling.
Each dictionary defines font, fill, alignment, border, and number_format
properties that map to openpyxl style objects.
"""

level_1_format = {
    'font': {
        'name': 'Palatino Linotype',
        'size': 12.0,
        'bold': True,
        'italic': True,
        'underline': None,
        'color': 'FF000000'
    },
    'fill': {
        'fill_type': None,
        'fgColor': '00000000',
        'bgColor': '00000000'
    },
    'alignment': {
        'horizontal': None,
        'vertical': None,
        'wrap_text': True,
        'indent': 0
    },
    'border': {
        'top': None,
        'bottom': None,
        'left': 'thin',
        'right': 'thin'
    },
    'number_format': 'General'
}

level_2_format = {
    'font': {
        'name': 'Palatino Linotype',
        'size': 11.0,
        'bold': True,
        'italic': False,
        'underline': None,
        'color': 'FF000000'
    },
    'fill': {
        'fill_type': None,
        'fgColor': '00000000',
        'bgColor': '00000000'
    },
    'alignment': {
        'horizontal': 'left',
        'vertical': None,
        'wrap_text': True,
        'indent': 1
    },
    'border': {
        'top': None,
        'bottom': None,
        'left': 'thin',
        'right': 'thin'
    },
    'number_format': 'General'
}

level_3_format = {
    'A': {
        'font': {'name': 'Palatino Linotype', 'size': 10.0, 'bold': False, 'italic': False, 'underline': None, 'color': 'FF000000'},
        'fill': {'fill_type': None, 'fgColor': '00000000', 'bgColor': '00000000'},
        'alignment': {'horizontal': 'left', 'vertical': None, 'wrap_text': False, 'indent': 2},
        'border': {'top': None, 'bottom': None, 'left': 'thin', 'right': 'thin'},
        'number_format': 'General'
    },
    'B': {
        'font': {'name': 'Palatino Linotype', 'size': 10.0, 'bold': False, 'italic': False, 'underline': None, 'color': 'FF000000'},
        'fill': {'fill_type': None, 'fgColor': '00000000', 'bgColor': '00000000'},
        'alignment': {'horizontal': 'center', 'vertical': 'center', 'wrap_text': False, 'indent': 0},
        'border': {'top': None, 'bottom': None, 'left': 'thin', 'right': 'thin'},
        'number_format': 'General'
    },
    'C': {
        'font': {'name': 'Palatino Linotype', 'size': 10.0, 'bold': False, 'italic': False, 'underline': None, 'color': 'FF000000'},
        'fill': {'fill_type': None, 'fgColor': '00000000', 'bgColor': '00000000'},
        'alignment': {'horizontal': 'center', 'vertical': 'center', 'wrap_text': False, 'indent': 0},
        'border': {'top': None, 'bottom': None, 'left': 'thin', 'right': 'thin'},
        'number_format': 'General'
    },
    'D': {
        'font': {'name': 'Palatino Linotype', 'size': 10.0, 'bold': False, 'italic': False, 'underline': None, 'color': 'FF000000'},
        'fill': {'fill_type': None, 'fgColor': '00000000', 'bgColor': '00000000'},
        'alignment': {'horizontal': None, 'vertical': 'center', 'wrap_text': False, 'indent': 0},
        'border': {'top': None, 'bottom': None, 'left': 'thin', 'right': 'thin'},
        'number_format': '_("$"* #,##0.00_);_("$"* \\(#,##0.00\\);_("$"* "-"??_);_(@_)'
    },
    'E': {
        'font': {'name': 'Palatino Linotype', 'size': 10.0, 'bold': False, 'italic': False, 'underline': None, 'color': 'FF000000'},
        'fill': {'fill_type': None, 'fgColor': '00000000', 'bgColor': '00000000'},
        'alignment': {'horizontal': None, 'vertical': 'center', 'wrap_text': False, 'indent': 0},
        'border': {'top': None, 'bottom': None, 'left': 'thin', 'right': 'thin'},
        'number_format': '_("$"* #,##0.00_);_("$"* \\(#,##0.00\\);_("$"* "-"??_);_(@_)'
    }
}

page_sum_format = {
    'font': {'name': 'Palatino Linotype', 'size': 10.0, 'bold': False, 'italic': False, 'underline': None, 'color': 'FF000000'},
    'fill': {'fill_type': None, 'fgColor': '00000000', 'bgColor': '00000000'},
    'alignment': {'horizontal': 'center', 'vertical': 'center', 'wrap_text': False, 'indent': 0},
    'border': {'top': 'thin', 'bottom': 'double', 'left': 'thin', 'right': 'thin'},
    'number_format': 'General'
}

amount_sum_format = {
    'font': {'name': 'Palatino Linotype', 'size': 10.0, 'bold': False, 'italic': False, 'underline': None, 'color': 'FF000000'},
    'fill': {'fill_type': None, 'fgColor': '00000000', 'bgColor': '00000000'},
    'alignment': {'horizontal': None, 'vertical': None, 'wrap_text': False, 'indent': 0},
    'border': {'top': 'thin', 'bottom': 'double', 'left': 'thin', 'right': 'thin'},
    'number_format': '_("$"* #,##0.00_);_("$"* \\(#,##0.00\\);_("$"* "-"??_);_(@_)'
}

level_1_page_sum_format = {
    'font': {'name': 'Palatino Linotype', 'size': 11.0, 'bold': True, 'italic': False, 'underline': None, 'color': 'FF000000'},
    'fill': {'fill_type': None, 'fgColor': '00000000', 'bgColor': '00000000'},
    'alignment': {'horizontal': 'center', 'vertical': 'center', 'wrap_text': False, 'indent': 0},
    'border': {'top': 'thin', 'bottom': 'thin', 'left': 'thin', 'right': 'thin'},
    'number_format': 'General'
}

level_1_amount_sum_format = {
    'font': {'name': 'Palatino Linotype', 'size': 11.0, 'bold': True, 'italic': False, 'underline': None, 'color': 'FF000000'},
    'fill': {'fill_type': None, 'fgColor': '00000000', 'bgColor': '00000000'},
    'alignment': {'horizontal': None, 'vertical': None, 'wrap_text': False, 'indent': 0},
    'border': {'top': 'thin', 'bottom': 'thin', 'left': 'thin', 'right': 'thin'},
    'number_format': '_("$"* #,##0.00_);_("$"* \\(#,##0.00\\);_("$"* "-"??_);_(@_)'
}

level_1_description_sum_format = {
    'font': {'name': 'Palatino Linotype', 'size': 11.0, 'bold': True, 'italic': False, 'underline': None, 'color': 'FF000000'},
    'fill': {'fill_type': None, 'fgColor': '00000000', 'bgColor': '00000000'},
    'alignment': {'horizontal': 'center', 'vertical': None, 'wrap_text': False, 'indent': 0},
    'border': {'top': 'thin', 'bottom': 'thin', 'left': 'thin', 'right': 'thin'},
    'number_format': 'General'
}
