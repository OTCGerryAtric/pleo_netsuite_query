import pandas as pd
import numpy as np
import os

from funtions import extract_gl_code
from funtions import extract_gl_name

def cash_balance_function():
    # Change Directory
    directory = r'I:\Shared drives\FP&A\Month End\00 - Python Code\NetSuite Extract'
    os.chdir(directory)

    # List all files in the directory
    all_files = os.listdir(directory)
    matching_files = [filename for filename in all_files if "Cash Balance" in filename]

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

        # Remove Non Bank Accounts
        bank_client_accounts_index = df[df['Financial Row'] == 'Total - Collateral Accounts'].index
        if not bank_client_accounts_index.empty:
            index_to_remove = bank_client_accounts_index[0]
            df = df.iloc[:index_to_remove]

        # Insert Category
        df['Category'] = df['Financial Row'].shift(1)
        df['Category'] = np.where(df['Category'].isin(['Bank', 'Bank - Client Money Accounts', 'Collateral Accounts']), df['Category'], np.nan)

        for i, row in df.iterrows():
            if i > 0 and pd.isna(row['Category']):
                df.at[i, 'Category'] = df.at[i - 1, 'Category']

        # Cleaned Output
        df = df.loc[df['GL Name'].notna()]
        cols = df.columns.tolist()
        first_cols = ['Financial Row', 'GL Code', 'GL Name']
        remaining_cols = [col for col in cols if col not in first_cols]
        df = df[first_cols + remaining_cols]
        df.iloc[:, 3:6] = df.iloc[:, 3:6].apply(lambda x: x.str.replace('€', '').str.replace(',', '.').str.replace(' ', ''))

        # Convert columns 3 and onward to numeric
        numeric_columns = df.columns[3:6]
        df[numeric_columns] = df[numeric_columns].apply(pd.to_numeric, errors='coerce')

        # Store the dataframe in the dictionary with the filename as the key
        dataframes[df_name] = df

        # Save Detailed Output
        save_directory = r'I:\Shared drives\FP&A\Month End\00 - Python Code'
        os.chdir(save_directory)
        df.to_csv('Bank Account Detail.csv', index=False)

        return df

# Call the function to create and process the DataFrame and assign it to a variable
cash_balance = cash_balance_function()