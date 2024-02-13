from openpyxl import load_workbook


def convert_text_to_numbers_with_openpyxl(excel_path):
    # Load the workbook and select the active worksheet
    workbook = load_workbook(filename=excel_path)
    worksheet = workbook.active  
    # Assuming you want to work with the first sheet

    # Iterate through all rows and columns, 
    # converting text to numbers where applicable
    for row in worksheet.iter_rows():
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

    # Save the modified workbook
    # Consider saving to a new file to preserve the original
    modified_excel_path = "modified_" + excel_path
    workbook.save(filename=modified_excel_path)


# Example usage
convert_text_to_numbers_with_openpyxl("data.xlsx")
