from col_names import col_names  # a list of columns to delete
from datetime import datetime
from delete_columns import delete_columns
from names_to_indices import names_to_indices
from openpyxl import load_workbook
from regex_replace import regex_replace
from resize_table import resize_table
from text_to_numbers import text_to_numbers

file_path = 'data.xlsx'
sheet_name = 'Sheet1'
table_name = 'Table1'
row_to_search = 1 

# Load the workbook
wb = load_workbook(file_path)
    
# Specify the sheet to work on
ws = wb[sheet_name]

# exports come with two extra rows at the top; delete them
ws.delete_rows(1, 2)

indices = names_to_indices(ws, col_names, row_to_search)

indices_list = sorted(indices.values(), reverse=True)
# print(indices_list)

delete_columns(ws, indices_list)

text_to_numbers(ws)

resize_table(ws, table_name)

regex_replace(ws, r'Curriculum\.', '')

# Save the modified workbook with a new name

current_date = datetime.now()
 
modified_file_path = current_date.strftime('%b %d %y') + ' BI ' + file_path
wb.save(filename=modified_file_path)
