from col_names import col_names # a list of columns to delete
from delete_columns import delete_columns
from names_to_indices import names_to_indices
from openpyxl import load_workbook
from text_to_numbers import convert_text_to_numbers


file_path = 'data.xlsx'
sheet_name = 'Sheet1'
row_to_search = 1

# Load the workbook
wb = load_workbook(file_path)
    
# Select the active worksheet or specify the sheet name
ws = wb.active  # or wb['SheetName']

# print(ws.tables.items())

indices = names_to_indices(ws, col_names, row_to_search)

indices_list = sorted(indices.values(), reverse=True)
# print(indices_list)

delete_columns(ws, indices_list)

# removes table formatting, since otherwise deleting columns gives a weird error from a damaged table
del ws.tables["Table1"]

convert_text_to_numbers(ws)

# Determine the used range
min_row = ws.min_row
max_row = ws.max_row
min_col = ws.min_column
max_col = ws.max_column

# Print the used range
print(f"Used range: Rows {min_row} to {max_row}, Columns {min_col} to {max_col}")

# Save the modified workbook
# Consider saving to a new file to preserve the original
modified_file_path = 'modified_' + file_path
wb.save(filename=modified_file_path)