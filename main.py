import pandas as pd
import openpyxl
import os

# Import Functions
from accounting_period import accounting_period_function
from cash_balance_eur import cash_balance_function
from tx_detail_eur import tx_detail_function

# Call the function to create and process the DataFrames and assign them to variables
acc_period = accounting_period_function()
tx_detail = tx_detail_function()
cash_balance = cash_balance_function()