def get_locations(tbl,
                  start_row,
                  start_col):
    """Provides row and column indices for the table elements.

    Args:
        tbl (tablespam.tablespam): table created with tablespam
        start_row (int): row index at which the table will start
        start_col (int): column index at which the table will start

    Returns:
        dict: indices of start and end rows and columns of all table elements 
    """

    # Initialize row indices
    if tbl.get('title') is not None:
        start_row_title = end_row_title = start_row
        start_row += 1
    else:
        start_row_title = end_row_title = None

    if tbl.get('subtitle') is not None:
        start_row_subtitle = end_row_subtitle = start_row
        start_row += 1
    else:
        start_row_subtitle = end_row_subtitle = None

    start_row_header = start_row

    if tbl['header'].get('lhs') is not None:
        start_row += max(tbl['header']['lhs']['level'], tbl['header']['rhs']['level']) - 1
    else:
        start_row += tbl['header']['rhs']['level'] - 1

    # Previous iteration was last entry of header rows, moving one up is the last row that belongs to the header
    end_row_header = start_row - 1

    # Begin data block at current row (one row below header)
    start_row_data = start_row
    end_row_data = start_row + len(tbl['table_data']['col_data']) - 1

    # Add one row for every entry in the table
    start_row += len(tbl['table_data']['col_data'])

    # Footnote is only one row long
    start_row_footnote = end_row_footnote = start_row

    # Initialize column indices
    start_col_title = start_col_subtitle = start_col_header_lhs = start_col_footnote = start_col
    n_col = 0

    # Get number of columns from `width` key in table header and add to start column
    # Skip lhs if formula had no lhs  
    if tbl['header'].get('lhs') is not None:
        n_col += tbl['header']['lhs']['width'] + tbl['header']['rhs']['width'] - 1
        start_col_header_rhs = start_col + tbl['header']['lhs']['width']
        end_col_header_lhs = start_col + tbl['header']['lhs']['width'] - 1
    else:
        n_col += tbl['header']['rhs']['width'] - 1
        start_col_header_rhs = start_col
        end_col_header_lhs = None
    end_col_title = end_col_subtitle = end_col_footnote = end_col_header_rhs = start_col + n_col

    # Return dictionary of all important indices
    return {
        'row': {
            'start_row_title': start_row_title,
            'end_row_title': end_row_title,
            'start_row_subtitle': start_row_subtitle,
            'end_row_subtitle': end_row_subtitle,
            'start_row_header': start_row_header,
            'end_row_header': end_row_header,
            'start_row_data': start_row_data,
            'end_row_data': end_row_data,
            'start_row_footnote': start_row_footnote,
            'end_row_footnote': end_row_footnote
        },
        'col': {
            'start_col_title': start_col_title,
            'end_col_title': end_col_title,
            'start_col_subtitle': start_col_subtitle,
            'end_col_subtitle': end_col_subtitle,
            'start_col_header_lhs': start_col_header_lhs,
            'end_col_header_lhs': end_col_header_lhs,
            'start_col_header_rhs': start_col_header_rhs,
            'end_col_header_rhs': end_col_header_rhs,
            'start_col_footnote': start_col_footnote,
            'end_col_footnote': end_col_footnote
        }
    }


def add_header_width(parsed_partial):
    """
    Adds width information to each element in a table header.

    Tablespan represents headers as nested lists. This function calculates how
    wide each header entry must be by determining how many root elements each
    parent element spans. For example, if a header 'x' spans two elements 'x1' and 'x2':

        |    x    |
        | x1 | x2 |

    The function updates each header entry with its corresponding width.

    Args:
        parsed_partial (dict): The left-hand side or right-hand side of the parsed table.

    Returns:
        dict or None: The parsed_partial with additional width fields, or None if parsed_partial is None.
    """
    if parsed_partial is None:
        return None

    # In case of single level set width to 1 (current parsed_partial is not a spanner) 
    if 'entries' not in parsed_partial or parsed_partial['entries'] is None:
        parsed_partial['width'] = 1
        return parsed_partial

    # Recursively go through all nested spanners and add to width
    parsed_partial['width'] = 0
    for entry_index in range(len(parsed_partial['entries'])):
        parsed_partial['entries'][entry_index] = add_header_width(parsed_partial['entries'][entry_index])
        parsed_partial['width'] += parsed_partial['entries'][entry_index]['width']

    return parsed_partial
