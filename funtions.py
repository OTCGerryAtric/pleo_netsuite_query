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