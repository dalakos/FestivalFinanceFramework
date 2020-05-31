# FestivalFinanceFramework
Main Festival Finance Code and Generation

This set of files are used to generate the main program and entire set of supporting files. The set are collectively referred to as the finance framework. The framework consists of the following:
1. ledger_XXXX, main accounting ledger for transactions before and after the festival event.
2. ledgerFestival_XXXX, accounting ledger for all transactions during the festival.
3. MAINXXX, main executable program.
4. CapComLedger -- main downloaded csv file of all transactions in main CapCom checking account.
5. CapComLedgerSavings -- Capcom savings account downloaded transactions.
6. festival_historical.xlsx is historical festival numbers generated from last year's festival and adding current year on to it. It is modified anytime main program is read.

SETUP
Pre-work: Fill out all accounts for the festival in the MasterStartFile.xlsx. It is strongly advised to keep all accounts even if they are not used.
Run the main executable file, RunMeFirst.py. This will generate the necessary files under the folder titled by YEAR. 
