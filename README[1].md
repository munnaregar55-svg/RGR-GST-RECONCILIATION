# RGR GST Reconciliation Web App

## Run on Windows
1. Install Python 3.10+.
2. Open Command Prompt in this folder.
3. Run: `pip install -r requirements.txt`
4. Run: `python app.py`
5. Open: `http://127.0.0.1:5000`

## What it does
- Books Excel + GSTR-2B Excel upload
- Auto-detects common GST columns
- Matches GSTIN + Invoice No
- ₹1 tolerance for taxable/IGST/CGST/SGST/total differences
- Shows Matched, Mismatch, Not in 2B, Not in Books and RC counts
- Exports exception report to Excel

For production use, add authentication, database, duplicate-invoice handling, multi-company support and advanced GSTR-2B formats.
