import pandas as pd
import openpyxl
import os

# Import Functions
from accounting_period import accounting_period_function
from subsiduary_eur import subsiduary_function
from domain_eur import domain_function
from competence_eur import competence_function
from marketing_eur import marketing_function
from cash_balance_eur import cash_balance_function
from tx_detail_eur import tx_detail_function
from funtions import autofit_columns_in_workbook
from funtions import freeze_first_row
from funtions import apply_custom_number_format
from funtions import apply_custom_date_format
from funtions import convert_text_to_date

# Call the function to create and process the DataFrames and assign them to variables
acc_period = accounting_period_function()
subsiduary = subsiduary_function()
domain = domain_function()
competence = competence_function()
marketing_raw = marketing_function()
tx_detail = tx_detail_function()
cash_balance = cash_balance_function()

# Set Shared Sheet Name
shared_sheet_name = ['Accounting Period', 'Subsiduary', 'Domain', 'Competence', 'Marketing']

# Set Output Directory
output_directory = r'I:\Shared drives\FP&A\Month End\00 - Python Code'
os.chdir(output_directory)

# Load the existing Excel workbook
ns_workbook_filename = 'NS Consolidation.xlsx'
ns_workbook = openpyxl.load_workbook(ns_workbook_filename)

# Create an ExcelWriter with the openpyxl engine using the existing workbook
with pd.ExcelWriter(ns_workbook_filename, engine='openpyxl') as writer:
    # Write each DataFrame to a separate sheet
    acc_period.to_excel(writer, sheet_name='Accounting Period', index=False)
    subsiduary.to_excel(writer, sheet_name='Subsiduary', index=False)
    domain.to_excel(writer, sheet_name='Domain', index=False)
    competence.to_excel(writer, sheet_name='Competence', index=False)
    marketing_raw.to_excel(writer, sheet_name='Marketing', index=False)

    # Get the workbook object
    workbook = writer.book

# Set Output Directory
output_directory = r'I:\Shared drives\FP&A\Month End\00 - Python Code'
os.chdir(output_directory)
#
# Save the updated Excel file
workbook.save('NS Consolidation.xlsx')

# Set Sheets to Format
shared_sheet_names = ['Accounting Period', 'Subsiduary', 'Domain', 'Competence', 'Marketing']

# Set Files
ns_workbook_filename = 'NS Consolidation.xlsx'

# Apply Functions
ns_workbook = openpyxl.load_workbook(ns_workbook_filename)
autofit_columns_in_workbook(ns_workbook, shared_sheet_names)
freeze_first_row(ns_workbook, shared_sheet_names)

# Apply custom number format to columns from F to Z in the specified sheets
custom_number_format = '#,##0.00_-;[Red](#,##0.00)_-;-_-'
apply_custom_number_format(ns_workbook, shared_sheet_names, 'F', 'Z', custom_number_format)

# Assuming ns_workbook is already loaded
if 'Accounting Period' in ns_workbook.sheetnames:
    sheet = ns_workbook['Accounting Period']
    for column in sheet.iter_cols(min_col=7, max_col=26, min_row=1, max_row=1):
        for cell in column:
            convert_text_to_date(cell)

# Apply the custom date format to columns F to Z in the 'Marketing' sheet
apply_custom_date_format(ns_workbook['Marketing'], 'F', 'Z', 'dd-mmm-yy')

# Save the updated workbooks
ns_workbook.save(ns_workbook_filename)

# Set Source and Destination Workbooks
source_workbook = openpyxl.load_workbook('NS Consolidation.xlsx')
destination_workbook = openpyxl.load_workbook('Month End Pack v1.0.xlsx')

for sheet_name in shared_sheet_name:
    # Clear existing data in the destination sheet
    sheet_to_copy = source_workbook[sheet_name]
    destination_sheet_name = sheet_name
    destination_sheet = destination_workbook[destination_sheet_name]
    destination_sheet.delete_rows(1, destination_sheet.max_row)

    # Copy data from the source sheet to the destination sheet as values
    for row in sheet_to_copy.iter_rows(values_only=True):
        destination_sheet.append(row)

# Save File
destination_workbook.save('Month End Pack v1.0.xlsx')

# Close the workbooks
source_workbook.close()
destination_workbook.close()

# Set Files
month_end_workbook_filename = 'Month End Pack v1.0.xlsx'

# Apply Functions
month_end_workbook = openpyxl.load_workbook(month_end_workbook_filename)
autofit_columns_in_workbook(month_end_workbook, shared_sheet_names)
freeze_first_row(month_end_workbook, shared_sheet_names)

# Apply custom number format to columns from F to Z in the specified sheets
custom_number_format = '#,##0.00_-;[Red](#,##0.00)_-;-_-'
apply_custom_number_format(month_end_workbook, shared_sheet_names, 'F', 'Z', custom_number_format)

# Apply the custom date format to columns F to Z in the 'Marketing' sheet
apply_custom_date_format(month_end_workbook['Accounting Period'], 'F', 'Z', 'dd-mmm-yy')
apply_custom_date_format(month_end_workbook['Marketing'], 'F', 'Z', 'dd-mmm-yy')

# Save the updated workbooks
month_end_workbook.save(month_end_workbook_filename)

# Print Missing GL's
missing_gls = acc_period[['Financial Row', 'GL Code', 'GL Name', 'GL Category', 'GL Group']]
missing_gls = missing_gls.loc[missing_gls['GL Category'].isna()]
missing_gls.to_csv("Missing GL's.csv", index=False)