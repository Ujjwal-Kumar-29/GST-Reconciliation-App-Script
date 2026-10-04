# GST Reconciliation (Purchase Register vs GSTR-2B)

A Google Apps Script tool that automatically matches purchase
register entries with GSTR-2B data and highlights differences.

## Features
- Creates a Google Sheet with purchase_register, GSTR2B and result tabs
- Matches invoices on GSTIN + invoice number, even when formats differ (INV/001 vs INV-1)
- Flags: matched, tax amount mismatch, in books only (ITC at risk), in 2B only (not booked)
- Colour-coded results with a summary of ITC at risk

## How to use
1. Open script.google.com and create a new project
2. Paste the code from GST_reconciliation.gs
3. Run setupGSTReconciliation and allow permissions
4. Open the generated sheet and check the result tab

## Note
All data in this project is dummy data. No client data is used.

## Built with
Google Apps Script, Google Sheets


### Here is a visual of final result:-
<img width="1531" height="647" alt="Screenshot 2026-10-04 162808" src="https://github.com/user-attachments/assets/acb95fc4-7043-4b62-8ed2-6f10ea63007e" />

