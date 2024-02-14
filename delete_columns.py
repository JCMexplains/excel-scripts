
from openpyxl import load_workbook


def delete_columns_in_excel(file_path, cols_to_delete):
    '''
    Delete specified columns from an Excel file using openpyxl.

    Parameters:
    - file_path: The path to the Excel file.
    - cols_to_delete: A list of column indexes to delete, 1-based indexing.
                      Must be sorted in descending order.
    '''
    # Load the workbook
    wb = load_workbook(file_path)
    
    # Select the active worksheet or specify the sheet name
    ws = wb.active  # or wb['SheetName']

    # Delete columns; iterate in reverse to avoid index shifting issues
    for col_index in sorted(cols_to_delete, reverse=True):
        ws.delete_cols(col_index, 1)

    # Save the modified workbook
    # Consider saving to a new file to preserve the original
    modified_file_path = 'modified_' + file_path
    wb.save(filename=modified_file_path)


# # Example usage:
# # ensure the list is 
# # in descending order for accurate deletion
# delete_columns_in_excel('data.xlsx', [4, 2])
