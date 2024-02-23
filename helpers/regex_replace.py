import re


def regex_replace(ws, find, replace):

    # Compile the regex pattern for efficiency
    pattern = re.compile(fr'{find}')

    # Iterate over all rows and columns
    for row in ws.iter_rows():
        for cell in row:
            if cell.value and isinstance(cell.value, str):
                # If the cell contains text, search for the pattern
                if pattern.search(cell.value):
                    # Replace the found text with the replacement text
                    cell.value = pattern.sub(replace, cell.value)
