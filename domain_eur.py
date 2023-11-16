import pandas as pd
import numpy as np
import os

from funtions import extract_gl_code
from funtions import extract_gl_name

def domain_function():
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
    matching_files = [filename for filename in all_files if "Domain" in filename]

    # Initialize an empty dictionary to store dataframes
    dataframes = {}

    # Loop through matching files and create dataframes
    for filename in matching_files:
        file_path = os.path.join(directory, filename)
        df = pd.read_csv(file_path, skiprows=6)
        df.columns = df.columns.str.strip()

        # Remove the file extension from the dataframe name
        df_name = os.path.splitext(filename)[0]

        # Apply the custom function to create the 'GL Code' column
        df['GL Code'] = df['Financial Row'].apply(extract_gl_code)
        df['GL Name'] = df['Financial Row'].apply(extract_gl_name)
        df['GL Name'] = np.where(df['GL Code'].str.isdigit(), df['GL Name'], np.nan)

        # Exclude Transfer Pricing
        df = df[~df['Financial Row'].str.contains('Transfer Pricing', case=False)]
        df = df.loc[df['GL Name'].notna()]

        # Apply Mapping Data
        df = pd.merge(df, gl_mapping_data[['GL Code', 'GL Category', 'GL Group', 'Signage']], on='GL Code', how='left')

        # Cleaned Output
        cols = df.columns.tolist()
        first_cols = ['Financial Row', 'GL Code', 'GL Name', 'GL Category', 'GL Group', 'Signage']
        remaining_cols = [col for col in cols if col not in first_cols]
        df = df[first_cols + remaining_cols]
        df = df.drop('Total', axis=1)
        df.iloc[:, 6:] = df.iloc[:, 6:].apply(
            lambda x: x.str.replace('€', '').str.replace(',', '.').str.replace(' ', ''))

        # Convert columns 4 and onward to numeric
        numeric_columns = df.columns[6:]
        df[numeric_columns] = df[numeric_columns].apply(pd.to_numeric, errors='coerce')

        # Apply the function to change signage to the numeric columns
        for col in numeric_columns:
            df[col] = df[col] * df['Signage']

        # Drop Switch Signage
        df = df.drop(columns='Signage')

        # Create GL Category Index
        df['Running Count'] = df.groupby('GL Category').cumcount() + 1
        df['GL Category Index'] = df['GL Category'] + ' ' + df['Running Count'].astype(str)
        df = df.drop(columns=['Running Count'])

        # Finalise Output
        cols = df.columns.tolist()
        first_cols = ['Financial Row', 'GL Code', 'GL Name', 'GL Category', 'GL Category Index', 'GL Group']
        remaining_cols = [col for col in cols if col not in first_cols]
        df = df[first_cols + remaining_cols]

        # Store the dataframe in the dictionary with the filename as the key
        dataframes[df_name] = df

        return df

# Call the function to create and process the DataFrame and assign it to a variable
domain = domain_function()