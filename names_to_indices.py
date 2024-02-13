from column_names import column_names
from openpyxl import load_workbook


def column_names_to_indices(file_path, sheet_name, column_names):
    """
    Convert column names to indices in an Excel sheet.

    Parameters:
    - file_path: Path to the Excel file.
    - sheet_name: Name of the worksheet to search for column names.
    - column_names: A list of column names for which indices are required.

    Returns:
    A dictionary with column names as keys and their 1-based indices as values.
    """
    # Load the workbook and select the specified worksheet
    wb = load_workbook(file_path)
    ws = wb[sheet_name]

    # Initialize a dictionary to hold the column names and their indices
    col_indices = {}

    # Iterate over the THIRD row to find columns and their indices
    for col in ws.iter_rows(min_row=3, max_row=3, values_only=True):
        for idx, cell_value in enumerate(col, start=1):  
            # start=1 for 1-based indexing
            if cell_value in column_names:
                col_indices[cell_value] = idx

    return col_indices


# Example usage
file_path = 'data.xlsx'
sheet_name = 'Sheet1'  # Adjust as necessary
indices = column_names_to_indices(file_path, sheet_name, column_names)
print(indices)
indices_list = list(indices.values())
indices_list.sort(reverse=True)
print(indices_list)
