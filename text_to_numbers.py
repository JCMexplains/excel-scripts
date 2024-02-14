
def convert_text_to_numbers(ws):
    # Takes an Excel worksheet from openpyxl as an argument

    # Iterate through all rows and columns, 
    # converting text to numbers where applicable
    for row in ws.iter_rows():
        for cell in row:
            if cell.data_type == 's':  
                # Check if the cell data type is string ('s')
                try:
                    # Attempt to convert the string to a float
                    cell.value = float(cell.value)
                    cell.data_type = 'n'  # Set cell type to numeric ('n')
                except ValueError:
                    # If conversion fails, leave the value as is
                    continue

# Example usage
# convert_text_to_numbers('data.xlsx')
