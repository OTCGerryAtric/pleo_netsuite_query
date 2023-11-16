import pandas as pd
import numpy as np
import os

from funtions import extract_gl_code
from funtions import extract_gl_name

def tx_detail_function():
    # Importing Mapping Data
    root_directory = r'I:\Shared drives\FP&A\Month End\00 - Python Code\Mapping Files'
    os.chdir(root_directory)
    gl_mapping_data = pd.read_csv('GL Mapping File.csv')
    gl_mapping_data['GL Code'] = gl_mapping_data['GL Code'].astype(str)

    # Change Directory
    directory = r'I:\Shared drives\FP&A\Month End\00 - Python Code\NetSuite Extract'
    os.chdir(directory)

    # List all files in the directory
    all_files = os.listdir(directory)
    matching_files = [filename for filename in all_files if "Tx Detail" in filename]

    # Initialize an empty dictionary to store dataframes
    dataframes = {}

    # Loop through matching files and create dataframes
    for filename in matching_files:
        file_path = os.path.join(directory, filename)
        df = pd.read_csv(file_path)
        df.columns = df.columns.str.strip()

        # Apply the custom function to create the 'GL Code' column
        df['GL Code'] = df['Account'].str.slice(0, 6)
        df['GL Name'] = df['Account'].str.slice(7, 1000)
        df['Financial Row'] = df['GL Code'] + ' - ' + df['GL Name']
        df['Month'] = df['Period']
        df['Subsidiary'] = df['Subsidiary'].str.split(':').str.get(-1).str.strip()

        # Filter to P&L
        mask_notna = pd.notna(df['GL Code'])
        df = df[mask_notna]
        mask = df['GL Code'].str.startswith(('4', '5', '6', '7', '8'))
        df = df[mask]

        # Filter and Re-Order Columns
        df = df[['Financial Row', 'Month', 'GL Code', 'GL Name', 'Subsidiary', 'Domain', 'Competence', 'Marketing expenses', 'Name', 'Date', 'Type', 'Transaction Number','Amount']]
        df = df.rename(columns={'Marketing expenses': 'Marketing Category'})

        # Apply Mapping Data
        df = pd.merge(df, gl_mapping_data[['GL Code', 'GL Category', 'GL Group', 'Tx Signage']], on='GL Code', how='left')

        # Apply the function to change signage to the numeric columns
        df['Amount'] = df['Amount'] * df['Tx Signage']

        # Drop Switch Signage
        df = df.drop(columns='Tx Signage')

        # Save Detailed Output
        save_directory = r'I:\Shared drives\FP&A\Month End\00 - Python Code'
        os.chdir(save_directory)
        df.to_csv('Tx Details.csv', index=False)

        # Create Pivot Table
        df['Subsidiary'] = df['Subsidiary'].fillna('Blank')
        df['Domain'] = df['Domain'].fillna('Blank')
        df['Competence'] = df['Competence'].fillna('Blank')
        df['Marketing Category'] = df['Marketing Category'].fillna('Unknown')
        pivot_df = df.groupby(['Month', 'GL Code', 'GL Name', 'Subsidiary', 'Domain', 'Competence', 'Marketing Category'])['Amount'].sum().reset_index()
        pivot_df['Marketing Category'] = np.where((pivot_df['Marketing Category'] != '620020') & (pivot_df['Marketing Category'] == 'Unknown'), '', pivot_df['Marketing Category'])
        pivot_df['Amount'] = pivot_df['Amount'].round(3)
        pivot_df = pivot_df.sort_values(by=['Month', 'GL Code'])
        pivot_df.to_csv('Tx Summary.csv', index=False)
    return df

# Call the function to create and process the DataFrame and assign it to a variable
tx_detail = tx_detail_function()