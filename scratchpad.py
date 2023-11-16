import pandas as pd
import numpy as np
import os
from datetime import datetime

from funtions import extract_gl_code
from funtions import extract_gl_name

def accounting_period_function():
    # Importing Mapping Data
    root_directory = r'I:\Shared drives\FP&A\Month End\00 - Python Code\Mapping Files'
    os.chdir(root_directory)
    gl_mapping_data = pd.read_csv('GL Mapping File.csv')
    gl_mapping_data['GL Code'] = gl_mapping_data['GL Code'].astype(str)

    # Change Directory
    root_directory = r'I:\Shared drives\FP&A\Month End\00 - Python Code\NetSuite Extract'
    os.chdir(root_directory)

    # Load Accounting Period, skipping the first 6 rows
    df = pd.read_csv(r'2023 - Accounting Period.csv', skiprows=6)
    df.columns = df.columns.str.strip()

    # Apply the custom function to create the 'GL Code' column
    df['GL Code'] = df['Financial Row'].apply(extract_gl_code)
    df['GL Name'] = df['Financial Row'].apply(extract_gl_name)
    df['GL Name'] = np.where(df['GL Code'].str.isdigit(), df['GL Name'], np.nan)

    # Exclude Transfer Pricing
    df = df[~df['Financial Row'].str.contains('Transfer Pricing', case=False)]
    df = df.loc[df['GL Name'].notna()]

    # Apply Mapping Data
    df = pd.merge(df, gl_mapping_data[['GL Code', 'GL Category', 'GL Group']], on='GL Code', how='left')
    df['GL Code'] = gl_mapping_data['GL Code'].astype(str)

    # Signage Switch
    df['Signage Switch'] = np.where(df['GL Code'].str[0].isin(['5', '6']), 'Yes', 'No')

    # Cleaned Output
    cols = df.columns.tolist()
    first_cols = ['Financial Row', 'GL Code', 'GL Name', 'GL Category', 'GL Group', 'Signage Switch']
    remaining_cols = [col for col in cols if col not in first_cols]
    df = df[first_cols + remaining_cols]
    df = df.drop('Total', axis=1)
    df.iloc[:, 6:] = df.iloc[:, 6:].apply(lambda x: x.str.replace('€', '').str.replace(',', '.').str.replace(' ', ''))

    # Convert relevant columns to numeric
    numeric_columns = df.columns[6:]
    df[numeric_columns] = df[numeric_columns].apply(pd.to_numeric, errors='coerce')

    # Apply the function to change signage to the numeric columns
    for col in numeric_columns:
        signage_multiplier = df['Signage Switch'].apply(lambda x: -1 if x == 'Yes' else 1)
        df[col] = df[col] * signage_multiplier

    # Drop Switch Signage
    df = df.drop(columns='Signage Switch')

    # Convert Column Headings to Dates
    column_headers = df.columns[5:].tolist()

    def convert_header_to_date(header):
        month_abbr, year_abbr = header.split(' ')
        month_number = datetime.strptime(month_abbr, '%b').strftime('%m')
        year = '20' + year_abbr  # Assuming the year format is 'yy'
        return f'01/{month_number}/{year}'

    # Use list comprehension to apply the conversion to selected column headers
    new_column_headers = [convert_header_to_date(header) for header in column_headers]

    # Update the selected column headers in the DataFrame
    df.columns = df.columns[:5].tolist() + new_column_headers

    return df

# Call the function to create and process the DataFrame and assign it to a variable
accounting_period = accounting_period_function()
