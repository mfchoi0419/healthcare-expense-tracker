from pathlib import Path
import re
import time
import shutil
import hashlib
import pandas as pd
import pdfplumber
from numbers_parser import Document

BASE_DIR = Path.home() / "Desktop" / "ExpensesTool"

INPUT_DIR = BASE_DIR / "input_files"
PROCESSED_DIR = BASE_DIR / "processed_files"
OUTPUT_DIR = BASE_DIR / "output"

MASTER_FILE = OUTPUT_DIR / "Healthcare_Expenses_Master.xlsx"
MERCHANT_FILE = BASE_DIR / "healthcare_merchants.csv"
LOG_FILE = OUTPUT_DIR / "processing_log.txt"

SUPPORTED_FILES = [".csv", ".xlsx", ".xls", ".pdf", ".numbers"]

DEFAULT_HEALTHCARE_MERCHANTS = [
    "MASS GENERAL",
    "MASS GEN",
    "BRIGHAM",
    "MGH",
    "ATRIUS",
    "CVS",
    "WALGREENS",
    "PHARMACY",
    "DENT",
    "DENTAL",
    "DERM",
    "DERMATOLOGY",
    "CHIRO",
    "CHIROPRACT",
    "ALLCARE MEDICAL",
    "MEDICAL",
    "HOSPITAL",
    "SURGI",
    "FOOT",
    "ANKLE",
    "ORTHOPEDIC",
    "JOHNSON COMPOUNDING",
    "BOYLSTON STREET DENT",
    "KRAUSS DERMATOLOGY",
    "MOVE WELL",
    "NEWTON WELLESLEY",
]


def log(message):
    OUTPUT_DIR.mkdir(exist_ok=True)
    with open(LOG_FILE, "a", encoding="utf-8") as f:
        f.write(f"{time.strftime('%Y-%m-%d %H:%M:%S')} - {message}\n")


def ensure_folders():
    INPUT_DIR.mkdir(exist_ok=True)
    PROCESSED_DIR.mkdir(exist_ok=True)
    OUTPUT_DIR.mkdir(exist_ok=True)

    if not MERCHANT_FILE.exists():
        pd.DataFrame({"merchant_name": DEFAULT_HEALTHCARE_MERCHANTS}).to_csv(
            MERCHANT_FILE,
            index=False
        )


def clean_amount(value):
    if pd.isna(value):
        return None

    text = str(value)
    text = text.replace("$", "").replace(",", "")
    text = text.replace("(", "-").replace(")", "")
    text = text.strip()

    try:
        return float(text)
    except ValueError:
        return None


def clean_company_name(value):
    text = str(value).strip()
    text = re.sub(r"\s+", " ", text)
    text = text.replace("AplPay ", "")
    return text


def load_healthcare_merchants():
    ensure_folders()
    df = pd.read_csv(MERCHANT_FILE)

    merchants = []
    for value in df["merchant_name"].dropna():
        merchants.append(str(value).upper().strip())

    return merchants


def is_healthcare_company(company_name):
    company_upper = str(company_name).upper()

    healthcare_merchants = load_healthcare_merchants()

    EXCLUDED_TERMS = [
        "VETERINARY",
        "VET",
        "VETCOVE",
        "ANIMAL",
        "HAMPSTEADANIMAL",
        "HAMPSTEAD ANIMAL",
        "PET",
        "DOG",
        "CAT",

        "NSTAR",
        "NATIONAL GRID",
        "SPEEDWAY",
        "GAS",
        "ELECTRIC",
        "UTILITY",
        "MARKET BASKET",
        "STOP & SHOP",
        "UBER",
        "LYFT",
        "SHELL",
        "MOBIL",
        "SUNOCO"
    ]

    if any(term in company_upper for term in EXCLUDED_TERMS):
        return False

    return any(
        merchant in company_upper
        for merchant in healthcare_merchants
    )


def get_row_value(row, possible_names):
    for name in possible_names:
        if name in row and pd.notna(row[name]):
            return row[name]
    return ""


def make_row_id(date, company, amount):
    raw = f"{date}|{company}|{amount}"
    return hashlib.sha256(raw.encode("utf-8")).hexdigest()[:16]


def process_csv_or_excel(file_path):
    if file_path.suffix.lower() == ".csv":
        df = pd.read_csv(file_path)
    else:
        df = pd.read_excel(file_path)

    rows = []

    for _, row in df.iterrows():
        company = get_row_value(row, [
            "Description",
            "Transaction",
            "Merchant",
            "Name",
            "Company",
            "Vendor / Charge Name"
        ])

        date = get_row_value(row, [
            "Date",
            "Transaction Date",
            "Posted Date"
        ])

        amount = get_row_value(row, [
            "Amount",
            "Charges",
            "Charge",
            "Price"
        ])

        company = clean_company_name(company)
        amount = clean_amount(amount)

        if not company or amount is None:
            continue

        if is_healthcare_company(company):
            rows.append({
                "Date": date,
                "Company": company,
                "Amount": amount
            })

    return pd.DataFrame(rows)


def parse_pdf_transaction_line(line):
    pattern = r"(\d{2}/\d{2}/\d{4})\s+([A-Za-z]+)\s+(.+?)\s+\$?(-?\d[\d,]*\.\d{2})$"
    match = re.search(pattern, line)

    if not match:
        return None

    return {
        "Date": match.group(1),
        "Company": clean_company_name(match.group(3)),
        "Amount": clean_amount(match.group(4))
    }


def process_pdf(file_path):
    rows = []

    inside_healthcare_section = False

    with pdfplumber.open(file_path) as pdf:
        for page in pdf.pages:
            text = page.extract_text() or ""

            for line in text.split("\n"):
                upper = line.upper()

                if "HEALTH CARE SERVICES" in upper or "PHARMACIES" in upper:
                    inside_healthcare_section = True
                    continue

                if (
                    "INSURANCE SERVICES" in upper
                    or "INTERNET SERVICES" in upper
                    or "RESTAURANT" in upper
                    or "TRANSPORTATION" in upper
                    or "TRAVEL" in upper
                    or "GROCERIES" in upper
                    or "MERCHANDISE" in upper
                ):
                    inside_healthcare_section = False

                parsed = parse_pdf_transaction_line(line)

                if not parsed:
                    continue

                company = parsed["Company"]

                if is_healthcare_company(company):
                    rows.append({
                        "Date": parsed["Date"],
                        "Company": company,
                        "Amount": parsed["Amount"]
                    })

    return pd.DataFrame(rows)


def append_to_master(new_data):
    if new_data.empty:
        return 0

    if MASTER_FILE.exists():
        existing = pd.read_excel(MASTER_FILE)
        combined = pd.concat([existing, new_data], ignore_index=True)
    else:
        combined = new_data

    combined = combined.drop_duplicates(
        subset=["Date", "Company", "Amount"],
        keep="first"
    )

    combined = combined.sort_values(by=["Date", "Company"], ascending=[True, True])

    combined.to_excel(MASTER_FILE, index=False)

    return len(new_data)

def process_numbers(file_path):
    doc = Document(str(file_path))
    rows = []

    for sheet in doc.sheets:
        for table in sheet.tables:
            data = table.rows(values_only=True)

            if not data or len(data) < 2:
                continue

            headers = [str(h).strip() if h else "" for h in data[0]]

            for row_values in data[1:]:
                row_dict = dict(zip(headers, row_values))

                company = (
                    row_dict.get("Description")
                    or row_dict.get("Transaction")
                    or row_dict.get("Merchant")
                    or row_dict.get("Name")
                    or row_dict.get("Company")
                    or ""
                )

                date = (
                    row_dict.get("Date")
                    or row_dict.get("Transaction Date")
                    or row_dict.get("Posted Date")
                    or ""
                )

                amount = (
                    row_dict.get("Amount")
                    or row_dict.get("Charges")
                    or row_dict.get("Charge")
                    or row_dict.get("Price")
                    or ""
                )

                company = clean_company_name(company)
                amount = clean_amount(amount)

                if not company or amount is None:
                    continue

                if is_healthcare_company(company):
                    rows.append({
                        "Date": date,
                        "Company": company,
                        "Amount": amount
                    })

    return pd.DataFrame(rows)

def process_file(file_path):
    log(f"Started processing: {file_path.name}")

    try:
        suffix = file_path.suffix.lower()

        if suffix in [".csv", ".xlsx", ".xls"]:
            data = process_csv_or_excel(file_path)
        elif suffix == ".pdf":
            data = process_pdf(file_path)
        elif suffix == ".numbers":
            data = process_numbers(file_path)
        else:
            log(f"Unsupported file skipped: {file_path.name}")
            return

        rows_found = append_to_master(data)

        destination = PROCESSED_DIR / file_path.name

        if destination.exists():
            destination = PROCESSED_DIR / f"{file_path.stem}_{int(time.time())}{file_path.suffix}"

        shutil.move(str(file_path), str(destination))

        log(f"Finished processing: {file_path.name}. Rows found: {rows_found}")

    except Exception as e:
        log(f"ERROR processing {file_path.name}: {e}")


def process_existing_files():
    files = sorted(INPUT_DIR.iterdir())

    if not files:
        log("No files found in input_files.")
        return

    for file_path in files:
        if file_path.suffix.lower() in SUPPORTED_FILES:
            process_file(file_path)


if __name__ == "__main__":
    ensure_folders()
    log("Healthcare Expense Tool started.")
    process_existing_files()
    log("Healthcare Expense Tool finished.")