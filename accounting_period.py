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
    df = pd.merge(df, gl_mapping_data[['GL Code', 'GL Class', 'GL Category', 'GL Group', 'Signage']], on='GL Code', how='left')
    df['GL Code'] = gl_mapping_data['GL Code'].astype(str)

    # Cleaned Output
    cols = df.columns.tolist()
    first_cols = ['Financial Row', 'GL Code', 'GL Name', 'GL Class', 'GL Category', 'GL Group', 'Signage']
    remaining_cols = [col for col in cols if col not in first_cols]
    df = df[first_cols + remaining_cols]
    df = df.drop('Total', axis=1)
    df.iloc[:, 7:] = df.iloc[:, 7:].apply(lambda x: x.str.replace('€', '').str.replace(',', '.').str.replace(' ', ''))

    # Convert relevant columns to numeric
    numeric_columns = df.columns[7:]
    df[numeric_columns] = df[numeric_columns].apply(pd.to_numeric, errors='coerce')

    # Apply the function to change signage to the numeric columns
    for col in numeric_columns:
        df[col] = df[col] * df['Signage']

    # Drop Switch Signage
    df = df.drop(columns='Signage')

    # Convert Column Headings to Dates
    column_headers = df.columns[5:].tolist()

    # Convert Column Headings to 'YYYY-MM-DD' format
    def convert_header_to_date(header):
        # Check if the header is a date in the 'DD/MM/YYYY' format
        if '/' in header:
            day, month, year = header.split('/')
            # Construct the date in 'YYYY-MM-DD' format
            formatted_date = f'{year}-{month}-{day}'
            return formatted_date
        else:
            return header

    # Apply the conversion to all column headers
    df.columns = [convert_header_to_date(header) for header in df.columns]

    # Create GL Category Index
    df['Running Count'] = df.groupby('GL Category').cumcount() + 1
    df['GL Category Index'] = df['GL Category'] + ' ' + df['Running Count'].astype(str)
    df = df.drop(columns=['Running Count'])

    # Check GL Number
    df['Check GL'] = df['Financial Row'].str[:7]
    df['GL Code'] = np.where(df['Check GL'] == df['GL Code'], df['GL Code'], df['Check GL'])

    # Finalise Output
    cols = df.columns.tolist()
    first_cols = ['Financial Row', 'GL Code', 'GL Name', 'GL Class', 'GL Category', 'GL Category Index', 'GL Group']
    remaining_cols = [col for col in cols if col not in first_cols]
    df = df[first_cols + remaining_cols]
    df = df.drop(columns=['Check GL'])

    # Save Output
    save_directory = r'I:\Shared drives\FP&A\Month End\00 - Python Code'
    os.chdir(save_directory)
    df.to_csv('Accounting Period.csv', index=False)

    # Save Missing GL's
    df = df.loc[df['GL Category'].isna()]
    df = df.iloc[:, :7]
    df = df.drop_duplicates()
    df.to_csv("Missing GL's.csv", index=False)

    return df

# Call the function to create and process the DataFrame and assign it to a variable
accounting_period = accounting_period_function()