import pandas as pd
import os

def marketing_function():
    # Change Directory
    root_directory = r'I:\Shared drives\FP&A\Month End\00 - Python Code\NetSuite Extract'
    os.chdir(root_directory)

    # Load Accounting Period, skipping the first 6 rows
    df = pd.read_csv(r'2023 - Marketing.csv', skiprows=6)
    df.columns = df.columns.str.strip()

    # Filter Columns
    df = df[['Subsidiary', 'Total', 'Accounting Period: Name', 'Domain: Name', 'Competence: Name', 'Operating Market', 'Marketing expenses: Name']]
    df = df.rename(columns={'Total': 'Amount', 'Accounting Period: Name': 'Month', 'Domain: Name': 'Domain', 'Competence: Name': 'Competence', 'Marketing expenses: Name': 'Marketing Category'})
    df['Subsidiary'] = df['Subsidiary'].str.split(':').str.get(-1)

    # Clean Data
    df['Amount'] = df['Amount'].str.replace('€', '').str.replace(',', '.').str.replace(' ', '')
    df = df.dropna(subset=['Subsidiary'])
    df['Amount'] = pd.to_numeric(df['Amount'])

    # Change Signage
    df['Amount'] = df['Amount'] * -1

    # Populate Missing Data with Unknown
    df[['Subsidiary', 'Domain', 'Competence', 'Operating Market', 'Marketing Category']] = df[['Subsidiary', 'Domain', 'Competence', 'Operating Market', 'Marketing Category']].fillna('Unknown')

    # Sort Columns
    df = df[['Subsidiary', 'Month', 'Domain', 'Competence', 'Operating Market', 'Marketing Category', 'Amount']]

    # Fix Dates
    df['Month'] = pd.to_datetime(df['Month'], format='%b %Y')

    # Pivot the DataFrame to get dates as columns and sum the "Amount" for each combination of "Month" and other columns
    pivot_df = df.pivot_table(index=["Subsidiary", "Domain", "Competence", "Operating Market", "Marketing Category"], columns="Month", values="Amount", aggfunc="sum", fill_value=0)

    # Reset the index to make columns from index levels
    pivot_df.reset_index(inplace=True)

    # Rename the index name for clarity
    pivot_df.index.name = None
    pivot_df = pivot_df.round(2)

    return pivot_df

marketing = marketing_function()
