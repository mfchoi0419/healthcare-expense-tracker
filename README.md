# Healthcare Expense Tracker

Automated healthcare expense extraction tool built in Python to assist my mother in filtering and organizing expenses.

This application processes:
- PDF credit card statements
- CSV exports
- Excel spreadsheets
- Apple Numbers files

and extracts likely healthcare-related expenses into a clean Excel summary.

## Features

- Automatic healthcare expense filtering
- PDF statement parsing
- Excel/CSV/Numbers support
- Duplicate prevention
- Merchant classification
- Veterinary expense exclusion
- Automatic Excel report generation
- macOS executable support

## Example Output

| Date | Company | Amount |
|------|------|------|
| XX/XX/XXXX | COMPANY | XX.XX |
| XX/XX/XXXX | COMPANY | XX.XX |

## Technologies Used

- Python
- pandas
- pdfplumber
- openpyxl
- numbers-parser
- PyInstaller

## Running Locally

Install dependencies:

```bash
pip install -r requirements.txt
```

Run:

```bash
python expense_watcher.py
```

## Packaging as macOS App

```bash
pyinstaller --onedir --name "Healthcare Expense Tool" expense_watcher.py
```

## Project Motivation

Built to simplify healthcare expense tracking for family financial organization and tax preparation.
