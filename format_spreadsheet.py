from datetime import datetime

import helpers
from openpyxl import load_workbook


def process_workbook(
        file_path, sheet_name, table_name, row_to_search, col_names):
    wb = load_workbook(file_path)
    ws = wb[sheet_name]

    ws.delete_rows(1, 2)  # Assuming first two rows are always to be deleted

    indices = helpers.names_to_indices(ws, col_names, row_to_search)
    indices_list = sorted(indices.values(), reverse=True)
    helpers.delete_columns(ws, indices_list)

    helpers.resize_table(ws, table_name)
    helpers.regex_replace(ws, r'Curriculum\.', '')  # delete this string
    helpers.regex_replace(ws, r'^(\d{3})0$', r'\1')  # trim trailing 0
    helpers.set_col_width(ws)
    # because regex searches work on text, 
    # best to keep the text_to_numbers call below the regex calls
    helpers.text_to_numbers(ws) 

    save_workbook_with_new_name(wb, file_path)


def save_workbook_with_new_name(wb, original_file_path):
    current_date = datetime.now().strftime('%b_%d_%Y')  
    modified_file_path = f'{current_date}_BI_{original_file_path}'
    wb.save(filename=modified_file_path)


if __name__ == '__main__':
    file_path = 'data.xlsx'
    sheet_name = 'Sheet1'
    table_name = 'Table1'
    row_to_search = 1
    from col_names import col_names  # import specific to this usage

    process_workbook(
        file_path, sheet_name, table_name, row_to_search, col_names
        )
