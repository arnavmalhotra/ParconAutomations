import openpyxl

def excel_to_dict(file_path, sheet_name=None):
    # Load the workbook
    wb = openpyxl.load_workbook(file_path, read_only=True, data_only=True)
    
    # Select the sheet
    if sheet_name:
        sheet = wb[sheet_name]
    else:
        sheet = wb.active
    
    # Get all rows, skipping the first three
    rows = list(sheet.iter_rows(min_row=2, values_only=True))
    
    # Assume the first row (4th row in Excel) contains headers
    headers = rows[0]
    
    # Create the dictionary of dictionaries
    result = {}
    for i, row in enumerate(rows[1:], start=1):  # Start enumeration from 1
        row_dict = {headers[j]: value for j, value in enumerate(row) if value is not None}
        if row_dict:  # Only add non-empty rows
            result[i] = row_dict
    
    return result

# Example usage
file_path = r'C:\Users\Arnav\Desktop\parcon\Infiniti\invoice_data.xlsx'
data = excel_to_dict(file_path)
print(data)