# Healthcare Expense Tracker

Automated healthcare expense extraction tool built in Python to assist with filtering and organizing real-world financial data from bank and credit card statements.

---

## Overview

This application processes financial documents and extracts healthcare-related transactions into a clean, structured Excel report.

It is designed to handle messy, inconsistent statement formats and automatically identify relevant expenses such as:

- Pharmacy purchases
- Medical visits and copays
- Hospital and clinic charges
- Vision and specialty care

---

## Features

- Robust PDF parsing (handles inconsistent bank statement formats)
- Automatic healthcare expense detection
- Fallback extraction for malformed lines
- Handles CR/DR financial formats (e.g., `59.48CR`)
- Duplicate prevention
- Filters out non-expense transactions (ATM, transfers, refunds)
- Supports:
  - PDF statements
  - CSV exports
  - Excel spreadsheets
  - Apple Numbers files
- Clean Excel output generation
- Packaged macOS executable (no Python required)

---

## How It Works

1. Extracts all transactions from input files  
2. Cleans and normalizes merchant names  
3. Filters out non-expense transactions  
4. Identifies healthcare-related vendors using keyword classification  
5. Outputs structured data into an Excel file  

---

## Example Output

| Date       | Company            | Amount |
|------------|--------------------|--------|
| 01/18/2025 | ATRIUS PHARMACY    | 23.30  |
| 04/12/2025 | CVS PHARMACY       | 20.71  |
| 07/05/2025 | PARTNERS           | 12.73  |

---

## Technologies Used

- Python  
- pandas  
- pdfplumber  
- openpyxl  
- numbers-parser  
- PyInstaller  

---

## Running Locally

Install dependencies:

```bash
pip install -r requirements.txt