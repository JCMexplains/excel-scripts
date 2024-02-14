from col_names import column_names # a list of columns to delete
from delete_columns import delete_columns_in_excel
from names_to_indices import column_names_to_indices
from openpyxl import load_workbook


file_path = 'data.xlsx'
sheet_name = 'Sheet1'
row_to_search = 1

# Load the workbook
wb = load_workbook(file_path)
    
# Select the active worksheet or specify the sheet name
ws = wb.active  # or wb['SheetName']

indices = column_names_to_indices(ws, column_names, row_to_search)

indices_list = sorted(indices.values(), reverse=True)
print(indices_list)

delete_columns_in_excel(ws, indices_list)

# Save the modified workbook
# Consider saving to a new file to preserve the original
modified_file_path = 'modified_' + file_path
wb.save(filename=modified_file_path)