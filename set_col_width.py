import openpyxl

def set_col_width(ws, columns=None, buffer=1.2):
    """
    Adjusts the width of columns in a ws to fit the text.
    
    Args:
    - ws: The ws to adjust column widths for.
    - columns: A list of column letters to adjust. Adjusts all columns if None.
    - buffer: A multiplier to apply to the width for a little extra space.
    """
    if columns is None:
        columns = [col[0].column_letter for col in ws.columns]
    
    for column in columns:
        max_length = 0
        for cell in ws[column]:
            try:
                # Adjust the column width if content is a string; consider buffer
                if len(str(cell.value)) > max_length:
                    max_length = len(cell.value)
            except:
                pass
        adjusted_width = (max_length + 2) * buffer  # Adding a default buffer for aesthetics
        ws.column_dimensions[column].width = adjusted_width

# # Example usage
# wb = openpyxl.load_workbook('your_workbook.xlsx')
# ws = wb.active

# set_col_width(ws, ['A', 'B'])  # Adjust columns A and B
# # Or, set_col_width(ws) to adjust all columns

# wb.save('adjusted_workbook.xlsx')
