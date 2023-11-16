import locale
import numpy as np
from datetime import datetime

# Define a custom function to split and check if it's a number
def extract_gl_code(row):
    parts = row.split(" - ", 1)
    if len(parts) > 1:
        gl_code = parts[0]
        # Check if gl_code is a number, if not, return a blank
        return gl_code if gl_code.isdigit() else ''
    else:
        return ''

def extract_gl_name(row):
    parts = row.split(" - ", 1)
    if len(parts) > 1:
        return parts[1]  # Return the part after the first " - "
    else:
        return ''

def format_numeric_columns(df):
    # Define the custom format
    custom_format = '#,##0.00_-;[Red](#,##0.00)_-;-_-'

    # Apply the custom format to all numeric columns
    numeric_columns = df.select_dtypes(include=[np.number]).columns
    df[numeric_columns] = df[numeric_columns].applymap(lambda x: locale.format_string(custom_format, x))
    return df

def autofit_columns_in_workbook(workbook, sheet_names):
    for sheet_name in sheet_names:
        sheet = workbook[sheet_name]
        for column_cells in sheet.columns:
            max_length = 0
            for cell in column_cells:
                try:
                    if len(str(cell.value)) > max_length:
                        max_length = len(str(cell.value))
                except:
                    pass
            adjusted_width = (max_length + 2) if (max_length + 2) > 15.71 else 15.71
            sheet.column_dimensions[column_cells[0].column_letter].width = adjusted_width

def freeze_first_row(workbook, sheet_names):
    for sheet_name in sheet_names:
        sheet = workbook[sheet_name]
        sheet.freeze_panes = 'A2'  # Freeze the first row (1-indexed)

def apply_custom_number_format(workbook, sheet_names, start_column, end_column, custom_format):
    for sheet_name in sheet_names:
        sheet = workbook[sheet_name]
        for col_letter in range(ord(start_column), ord(end_column) + 1):
            col_letter = chr(col_letter)
            for cell in sheet[col_letter]:
                cell.number_format = custom_format

def apply_custom_date_format(sheet, start_column, end_column, custom_format):
    for col_letter in range(ord(start_column), ord(end_column) + 1):
        col_letter = chr(col_letter)
        cell = sheet[f'{col_letter}1']  # Assuming the header is in row 1
        cell.number_format = custom_format

def convert_text_to_date(cell):
    if isinstance(cell.value, str):
        try:
            date_value = datetime.strptime(cell.value, "%d/%m/%Y")
            cell.value = date_value
            cell.number_format = 'DD-MMM-YY'  # Apply the desired date format
        except ValueError:
            pass