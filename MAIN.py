# -*- coding: utf-8 -*-
"""
Spyder Editor

"""

import numpy as np
import pandas as pd
import os
import time
from datetime import datetime, date
import babel.numbers
from babel.numbers import format_currency

today = date.today(); date_print = today.strftime("%B %d, %Y")

year = 2025

skip_row_num = 0 # use 3 for capcom

os.chdir("..")
cwd = os.getcwd()
#dir_path = os.path.dirname(os.path.realpath(__file__))

folder_name = str(year); file_name = 'ledger_' + str(year) + '.csv'; file_name_checking = 'Checking.csv'

path = '/' + folder_name + '/' + file_name
path_checking = '/' + folder_name + '/' + file_name_checking
path_savings = '/' + folder_name + '/Savings.csv'

# if file_name_checking == '':
#     path_checking = '/' + folder_name + '/Broadview_Checking.csv'

df = pd.DataFrame( pd.read_csv (cwd + path) )

df['Date'] = pd.to_datetime(df.date); df = df.sort_values(['Date'])

checking_file_creation_time = os.path.getctime(cwd + path_checking)
savings_file_creation_time = os.path.getctime(cwd + path_savings)

check_file_creation_str = time.ctime(checking_file_creation_time)
saving_file_creation_str = time.ctime(savings_file_creation_time)

check_file_creation = datetime.strptime(check_file_creation_str, '%a %b %d %H:%M:%S %Y')
saving_file_creation = datetime.strptime(saving_file_creation_str, '%a %b %d %H:%M:%S %Y')

ledger_last_date = df['Date'].iloc[-1]

checking_account_df = pd.DataFrame( pd.read_csv (cwd + path_checking , skiprows=skip_row_num))
checking_account_df['Date'] = checking_account_df['Posting Date']
checking_account_df['Date'] = pd.to_datetime(checking_account_df['Date'])
checking_account_df = checking_account_df.sort_values(by='Date')

savings_account_df = pd.DataFrame( pd.read_csv (cwd + path_savings , skiprows=skip_row_num))
savings_account_df['Date'] = savings_account_df['Posting Date']
savings_account_df['Date'] = pd.to_datetime(savings_account_df['Date'])
savings_account_df = savings_account_df.sort_values(by='Date')

# Check that the starting date does not have a transaction amount
start_date_transaction = 0
start_date_transaction_savings = 0


checking_account_df['Amount Credit'] = np.zeros(len(checking_account_df))
checking_account_df['Amount Debit'] = np.zeros(len(checking_account_df))

for i in range(len(checking_account_df)):
    if checking_account_df['Amount'][i] > 0: 
        checking_account_df.loc[i,'Amount Credit'] = checking_account_df.loc[i, 'Amount'] 
    else:
        checking_account_df.loc[i,'Amount Debit'] = checking_account_df.loc[i,'Amount']
        
checking_account_df['Amount Credit'].replace(0, '', inplace=True)
checking_account_df['Amount Debit'].replace(0, '', inplace=True)

if checking_account_df['Amount Credit'].iloc[0] != '':
    start_date_transaction = checking_account_df['Amount Credit'].iloc[0]
    
if savings_account_df['Amount'].iloc[0] != '':
    start_date_transaction_savings = savings_account_df['Amount'].iloc[0]    

checking_account_df['Dividend'] = np.zeros(len(checking_account_df))
savings_account_df['Dividend'] = np.zeros(len(savings_account_df))

for i in range(len(checking_account_df)):
    if ('Dividend' in checking_account_df['Description'][i]) | ('DIVIDEND' in checking_account_df['Description'][i]): 
        checking_account_df.loc[i,'Dividend'] = checking_account_df.loc[i, 'Amount']

for i in range(len(savings_account_df)):
    if ('Dividend' in savings_account_df['Description'][i]) | ('DIVIDEND' in savings_account_df['Description'][i]): 
        savings_account_df.loc[i,'Dividend'] = savings_account_df.loc[i, 'Amount']


dividends_checking = np.sum(checking_account_df['Dividend']) 
dividends_savings = np.sum(savings_account_df['Dividend']) 
dividends = dividends_checking + dividends_savings

start_date_transaction = round(start_date_transaction,2)
start_date = checking_account_df['Date'].iloc[0]
end_date = checking_account_df['Date'].iloc[-1]
start_balance = checking_account_df['Balance'].iloc[0]
end_balance = checking_account_df['Balance'].iloc[-1]

start_date_transaction_savings = round(start_date_transaction_savings,2)
start_date_savings = savings_account_df['Date'].iloc[0]
end_date_savings = savings_account_df['Date'].iloc[-1]
start_balance_savings = savings_account_df['Balance'].iloc[0]
end_balance_savings = savings_account_df['Balance'].iloc[-1]

warning_ledger_date_flag = "OK"

if (check_file_creation < ledger_last_date) & (saving_file_creation < ledger_last_date):
    warning_ledger_date_flag = "FAILED: LEDGER LAST ENTRY DATE COMES AFTER CHECKING ACCOUNT LAST DATE"

# Calculations
balance_checking = end_balance - start_balance
balance_savings = end_balance_savings - start_balance_savings
combined_start_balance = start_balance + start_balance_savings
change_accounts = end_balance - start_balance + end_balance_savings - start_balance_savings

options = ['credit', 'deposit','startup'] 
df_filt = df[np.isin(df['account'], options, invert=True)]
df_filt['account'] = df_filt['account'].replace({'Vasil PMT': 'vendor PMT', 'Wilson PMT': 'vendor PMT'})

df_recon = df[np.isin(df['account'], options, invert=False)]
df_recon = df_recon[df_recon['reconcile_flag'] == 1]
df_payables =  df[df['check_number'].notnull()]
payables_df =  df_payables[df_payables['payable_flag'] == 1]
payables = round(np.sum(payables_df['amount']),2)
receivables_df =  df[(df['pmt'] == 'income') & (df['receivable_flag'] == 1)]
receivables = round(np.sum(receivables_df['amount']),2)

df_income = df_filt.loc[df_filt['pmt']=='income']
df_expense = df_filt.loc[df_filt['pmt']=='expense']

df_monetary_expenses = df_expense[(df_expense['reconcile_flag'] == 1) & (df_expense['ledger_bank_reconcile_flag'] == 1) & (df_expense['payable_flag'] != 1) & (df_expense['receivable_flag'] != 1)]
df_non_monetary_expenses = df_expense[(df_expense['reconcile_flag'] == 1) & (df_expense['ledger_bank_reconcile_flag'] != 1) & (df_expense['payable_flag'] != 1) & (df_expense['receivable_flag'] != 1)]
df_accidental_expenses = df_expense[(df_expense['reconcile_flag'] != 1) & (df_expense['ledger_bank_reconcile_flag'] == 1) & (df_expense['payable_flag'] != 1) & (df_expense['receivable_flag'] != 1)]
df_payables_expenses = df_expense[(df_expense['reconcile_flag'] == 1) & (df_expense['ledger_bank_reconcile_flag'] != 1) & (df_expense['payable_flag'] == 1) & (df_expense['receivable_flag'] != 1)]
df_final_expenses = df_expense[(df_expense['reconcile_flag'] == 1)]

df_monetary_income = df_income[(df_income['reconcile_flag'] == 1) & (df_income['ledger_bank_reconcile_flag'] == 1) & (df_income['payable_flag'] != 1) & (df_income['receivable_flag'] != 1)]
df_non_monetary_income = df_income[(df_income['reconcile_flag'] == 1) & (df_income['ledger_bank_reconcile_flag'] != 1) & (df_income['payable_flag'] != 1) & (df_income['receivable_flag'] != 1)]
df_accidental_income = df_income[(df_income['reconcile_flag'] != 1) & (df_income['ledger_bank_reconcile_flag'] == 1) & (df_income['payable_flag'] != 1) & (df_income['receivable_flag'] != 1)]
df_receivables_income = df_income[(df_income['reconcile_flag'] == 1) & (df_income['ledger_bank_reconcile_flag'] != 1) & (df_income['payable_flag'] != 1) & (df_income['receivable_flag'] == 1)]
df_final_income = df_income[(df_income['reconcile_flag'] == 1)]

monetary_expenses = np.sum(df_monetary_expenses['amount'])
non_monetary_expenses = np.sum(df_non_monetary_expenses['amount'])
accidental_expenses = np.sum(df_accidental_expenses['amount'])
payables_expenses = np.sum(df_payables_expenses['amount'])

monetary_income = np.sum(df_monetary_income['amount'])
non_monetary_income = np.sum(df_non_monetary_income['amount'])
accidental_income = np.sum(df_accidental_income['amount'])
receivables_income = np.sum(df_receivables_income['amount'])

all_expenses = round(monetary_expenses + non_monetary_expenses + accidental_expenses,2)
final_expenses = round(all_expenses - accidental_expenses,2)

all_income = round(monetary_income + non_monetary_income + accidental_income,2)
final_income = round(all_income - accidental_income,2)

BANK_RECONCILE = all_income - all_expenses + payables_expenses - receivables_income + dividends
BANK_SUMMARY = (end_balance - start_balance) + (end_balance_savings - start_balance_savings)

#df_expense['account'] = df_expense['account'].replace({'$1 raffle' : 'raffle prizes','$100 raffle': 'raffle prizes'})

grp = df_filt.groupby(['account','pmt'])
grp_income = df_final_income.groupby(['account','pmt'])
grp_expense = df_final_expenses.groupby(['account','pmt'])
final_table_income = grp_income.agg({'amount' : 'sum', 'account' : 'count'})
final_table_expense = grp_expense.agg({'amount' : 'sum', 'account' : 'count'})

# generate copies of final tables for formatting
final_table_income_f = final_table_income
final_table_expense_f = final_table_expense

total_income = np.sum(final_table_income['amount'])
total_expense = np.sum(final_table_expense['amount'])
profit = total_income - total_expense

recon_records = profit + dividends - receivables + payables
recon_number = recon_records - ((end_balance - start_balance) + (end_balance_savings - start_balance_savings))


# Format cuurency values for printing and writing to report.                       
total_expense_f = babel.numbers.format_currency(total_expense, "USD", locale='en_US')
total_income_f = babel.numbers.format_currency(total_income, "USD", locale='en_US')
profit_f = babel.numbers.format_currency(total_income - total_expense, "USD", locale='en_US')

dividends_f = babel.numbers.format_currency(dividends, "USD", locale='en_US')
start_balance_f = babel.numbers.format_currency(start_balance, "USD", locale='en_US')
end_balance_f = babel.numbers.format_currency(end_balance, "USD", locale='en_US')
balance_checking_f = babel.numbers.format_currency(balance_checking, "USD", locale='en_US')
start_balance_savings_f = babel.numbers.format_currency(start_balance_savings, "USD", locale='en_US')
end_balance_savings_f = babel.numbers.format_currency(end_balance_savings, "USD", locale='en_US')
balance_savings_f = babel.numbers.format_currency(balance_savings, "USD", locale='en_US')
combined_start_balance_f = babel.numbers.format_currency(combined_start_balance, "USD", locale='en_US')
change_accounts_f = babel.numbers.format_currency(change_accounts, "USD", locale='en_US')
payables_f = babel.numbers.format_currency(payables, "USD", locale='en_US')
receivables_f = babel.numbers.format_currency(receivables, "USD", locale='en_US')

recon_records_f = babel.numbers.format_currency(recon_records, "USD", locale='en_US')
recon_number_f = babel.numbers.format_currency(recon_number, "USD", locale='en_US')

# Change table format to currency. Note that you CANNOT perform math on these columns now.
final_table_expense_f["amount"] = final_table_expense_f["amount"].apply(lambda x: format_currency(x, currency="USD", locale="en_US"))
final_table_income_f["amount"] = final_table_income_f["amount"].apply(lambda x: format_currency(x, currency="USD", locale="en_US"))

print("---------------------------------------------------")
print("          Festival finance SUMMARY (" + date_print + ')           ')
print("---------------------------------------------------")
print('Total income                  ',total_income_f)
print('Total expenses                ',total_expense_f)
print('Profit                        ', profit_f)
print("---------------------------------------------------")
print("          Broadview checking account balance")
print("---------------------------------------------------")
print('Transactions file created:    ', check_file_creation_str,')')
print('Starting checking balance:    ', start_balance_f, '(',start_date,')')
print('End balance:                  ', end_balance_f, '(',end_date,')')
print('Change in funds:              ', balance_checking_f)
print('# of transactions:            ', str(len(checking_account_df.index)))
print("---------------------------------------------------")
print("          Broadview savings account balance")
print("---------------------------------------------------")
print('Transactions file created:    ', saving_file_creation_str,')')
print('Starting savings balance:     ', start_balance_savings_f, '(',start_date_savings,')')
print('End balance:                  ', end_balance_savings_f, '(',end_date_savings,')')
print('Change in funds:              ', balance_savings_f)
print('# of transactions:            ', str(len(savings_account_df.index)))
print("---------------------------------------------------")
print("Checking and savings combined change in funds: ", change_accounts_f)
print("---------------------------------------------------")
print("          Reconcile festival numbers with checking and savings accounts")
print("---------------------------------------------------")
print("Start with profit:            ", profit_f)
print("add in dividends              + ", dividends_f)
print("add in payables               + ", payables_f)
print("subtract out receivables      - ", receivables_f)
print('Reconciled balance calc -->   ', recon_records_f)
print("---------------------------------------------------")


if (recon_number < 0.01) & (recon_number > -0.01):
    print("***Sum accounts account balance matches ledger balance (RECONCILED)")
elif recon_number > 0:
    print("!!!Sum accounts balance LOWER than accounted by " + recon_number_f)
else:
    print("!!!Sum accounts balance HIGHER than accounted by " + recon_number_f)

print("---------------------------------------------------")
print("")
print("---------------------------------------------------")
print("Income and expenses broken down by area")
print("---------------------------------------------------")
print("    Expenses")
print("---------------------------------------------------")
print(final_table_expense_f)
print("---------------------------------------------------")
print("    Income")
print("---------------------------------------------------")
print(final_table_income_f)
print("---------------------------------------------------")
print("---------------------------------------------------")
print("Data quality checks")
print("---------------------------------------------------")
print("Ledger check date: " + warning_ledger_date_flag)
print("---------------------------------------------------")
print("    ...end of report.")
print("---------------------------------------------------")


# Print out report to a file

f = open(folder_name + "/" + folder_name+ "_festival_report_" + date_print + ".txt", "w")
print("---------------------------------------------------", file=f)
print("          Festival finance SUMMARY (" + date_print + ')           ', file=f)
print("  Written by George Dalakos -- 518-419-9848, gdalakos@gmail.com", file=f)
print(" ", file=f)
print("  The festival finance numbers are compared against the Broadview savings and checking accounts", file=f)
print("  Any differences between the two are noted, otherwise, they are determined to be reconciled.", file=f)
print("  The festival finance numbers are also broken down by income and expenses into their respective areas.", file=f)
print("---------------------------------------------------", file=f)
print('Total income                  ',total_income_f, file=f)
print('Total expenses                ',total_expense_f, file=f)
print('Profit                        ',profit_f, file=f)
print("---------------------------------------------------", file=f)
print("          Broadview checking account balance", file=f)
print("---------------------------------------------------", file=f)
print('Transactions file created:    ', check_file_creation_str,')', file=f)
print('Starting checking balance:    ', start_balance_f, '(',start_date,')', file=f)
print('End balance:                  ', end_balance_f, '(',end_date,')', file=f)
print('Change in funds:              ', balance_checking_f , file=f)
print('# of transactions:            ', str(len(checking_account_df.index)), file=f)
print("---------------------------------------------------", file=f)
print("          Broadview savings account balance", file=f)
print("---------------------------------------------------", file=f)
print('Transactions file created:    ', saving_file_creation_str,')', file=f)
print('Starting savings balance:     ', start_balance_savings_f, '(',start_date_savings,')', file=f)
print('End balance:                  ', end_balance_savings_f, '(',end_date_savings,')', file=f)
print('Change in funds:              ', balance_savings_f , file=f)
print('# of transactions:            ', str(len(savings_account_df.index)), file=f)
print("---------------------------------------------------", file=f)
print("Checking and savings combined change in funds: ", change_accounts_f, file=f)
print("---------------------------------------------------", file=f)
print("          Reconcile festival numbers with checking and savings accounts", file=f)
print("---------------------------------------------------", file=f)
print("Start with profit:            ", profit_f, file=f)
print("add in dividends              + ", dividends_f, file=f)
print("add in payables               + ", payables_f, file=f)
print("subtract out receivables      - ", receivables_f, file=f)
print('Reconciled balance calc -->   ', recon_records_f, file=f)
print("---------------------------------------------------", file=f)

# Match up checking account and ledger balances and make assessment
if (recon_number < 0.01) & (recon_number > -0.01):
    print("***Sum accounts balance matches ledger balance (RECONCILED)", file=f)
elif recon_number > 0:
    print("!!!Sum accounts balance LOWER than accounted by " + recon_number_f, file=f)
else:
    print("!!!Sum accounts balance HIGHER than accounted by " + recon_number_f, file=f)
    
print("---------------------------------------------------", file=f)
print("", file=f)
print("---------------------------------------------------", file=f)
print("Income and expenses broken down by area", file=f)
print("---------------------------------------------------", file=f)
print("    Expenses", file=f)
print("---------------------------------------------------", file=f)
print(final_table_expense_f, file=f)
print("---------------------------------------------------", file=f)
print("    Income", file=f)
print("---------------------------------------------------", file=f)
print(final_table_income_f, file=f)
print("---------------------------------------------------", file=f)
print("---------------------------------------------------", file=f)
print("Data quality checks", file=f)
print("---------------------------------------------------", file=f)
print("Ledger check date: " + warning_ledger_date_flag, file=f)
print("---------------------------------------------------", file=f)
print("    ...end of report.", file=f)
print("---------------------------------------------------", file=f)

f.close()
