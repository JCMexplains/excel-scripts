


def delete_columns(ws, cols_to_delete):
    '''
    Delete specified columns from an Excel file using openpyxl.

    Parameters:
    - ws: Excel worksheet object
    - cols_to_delete: A list of column indexes to delete, 1-based indexing.
                      Must be sorted in descending order.
    '''


    # Delete columns; iterate in reverse to avoid index shifting issues
    for col_index in sorted(cols_to_delete, reverse=True):
        ws.delete_cols(col_index, 1)

# # Example usage:
# # ensure the list is 
# # in descending order for accurate deletion
# delete_columns('data.xlsx', [4, 2])
