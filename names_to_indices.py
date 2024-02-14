
def column_names_to_indices(ws, column_names, row_to_search):
    '''
    Convert column names to indices in an Excel sheet.

    Parameters:
    - ws: an Excel worksheet object
    - column_names: A list of column names for which indices are required.

    Returns:
    A dictionary with column names as keys and their 1-based indices as values.
    '''

    # Find column indices in the third row
    col_indices = {cell_value: idx for idx, cell_value in enumerate(
        ws[row_to_search], start=1) if cell_value.value in column_names}

    return col_indices

