from pathlib import Path
import re
import time
import shutil
import hashlib
import pandas as pd
import pdfplumber
import requests
from bs4 import BeautifulSoup
from numbers_parser import Document

# ==============================
# PATHS
# ==============================

BASE_DIR = Path.home() / "Desktop" / "ExpensesTool"

INPUT_DIR = BASE_DIR / "input_files"
PROCESSED_DIR = BASE_DIR / "processed_files"
OUTPUT_DIR = BASE_DIR / "output"

MASTER_FILE = OUTPUT_DIR / "Healthcare_Expenses_Master.xlsx"
MERCHANT_FILE = BASE_DIR / "healthcare_merchants.csv"
LOG_FILE = OUTPUT_DIR / "processing_log.txt"

SUPPORTED_FILES = [".csv", ".xlsx", ".xls", ".pdf", ".numbers"]

# ==============================
# DEFAULT MERCHANTS
# ==============================

DEFAULT_HEALTHCARE_MERCHANTS = [
    "MASS GENERAL", "BRIGHAM", "MGH",
    "ATRIUS", "CVS", "WALGREENS",
    "PHARMACY", "DENT", "DERM",
    "CHIROPRACT", "MEDICAL", "HOSPITAL",
    "SURGI", "ORTHOPEDIC",
    "JOHNSON COMPOUNDING",
    "BOYLSTON STREET DENT",
    "KRAUSS DERMATOLOGY",
    "MOVE WELL", "NEWTON WELLESLEY",
    "BLUE CROSS", "BCBS",
    "TRINET", "TRINET3",
    "HIGHEND CARE", "HIGH END CARE"
]

# ==============================
# LOGGING
# ==============================

def log(message):
    OUTPUT_DIR.mkdir(exist_ok=True)
    with open(LOG_FILE, "a", encoding="utf-8") as f:
        f.write(f"{time.strftime('%Y-%m-%d %H:%M:%S')} - {message}\n")

# ==============================
# SETUP
# ==============================

def ensure_folders():
    INPUT_DIR.mkdir(exist_ok=True)
    PROCESSED_DIR.mkdir(exist_ok=True)
    OUTPUT_DIR.mkdir(exist_ok=True)

    if not MERCHANT_FILE.exists():
        pd.DataFrame({"merchant_name": DEFAULT_HEALTHCARE_MERCHANTS}).to_csv(
            MERCHANT_FILE, index=False
        )

# ==============================
# CLEANING
# ==============================

def clean_amount(value):
    if pd.isna(value):
        return None

    text = str(value).replace("$", "").replace(",", "")
    text = text.replace("(", "-").replace(")", "").strip()

    try:
        return float(text)
    except ValueError:
        return None


def clean_company_name(value):
    text = str(value).strip()
    text = re.sub(r"\s+", " ", text)
    text = text.replace("AplPay ", "")
    return text

# ==============================
# MERCHANT LOGIC
# ==============================

def load_healthcare_merchants():
    df = pd.read_csv(MERCHANT_FILE)
    return [str(v).upper().strip() for v in df["merchant_name"].dropna()]


def check_online_healthcare(company_name):
    query = f"{company_name} healthcare medical insurance provider"
    url = "https://duckduckgo.com/html/"

    try:
        response = requests.post(
            url,
            data={"q": query},
            headers={"User-Agent": "Mozilla/5.0"},
            timeout=6
        )

        if response.status_code != 200:
            return False

        soup = BeautifulSoup(response.text, "html.parser")
        text = soup.get_text(" ", strip=True).lower()

        keywords = [
            "healthcare", "medical", "clinic",
            "hospital", "physician",
            "insurance", "pharmacy", "dental"
        ]

        return sum(k in text for k in keywords) >= 2

    except Exception:
        return False


def normalize(text):
    return re.sub(r'[^A-Z0-9]', '', str(text).upper())


def is_healthcare_company(company_name):
    company_raw = str(company_name).upper()
    company_norm = normalize(company_name)

    EXCLUDED_TERMS = [
        "VETERINARY", "VET", "ANIMAL",
        "NSTAR", "NATIONALGRID", "SPEEDWAY",
        "GAS", "ELECTRIC", "UTILITY",
        "MARKETBASKET", "STOPSHOP",
        "UBER", "LYFT", "SHELL", "MOBIL"
    ]

    # Normalize exclusions too
    if any(normalize(term) in company_norm for term in EXCLUDED_TERMS):
        return False

    merchants = load_healthcare_merchants()

    # Normalize merchant list too
    for m in merchants:
        if normalize(m) in company_norm:
            return True

    # fallback for tricky known patterns
    if any(x in company_norm for x in ["BCBS", "BLUECROSS", "TRINET", "HIGHENDCARE"]):
        return True

    # smart online fallback (only when needed)
    LIKELY = ["HEALTH", "CARE", "MED", "CLINIC", "RX", "INSURANCE"]

    if any(w in company_raw for w in LIKELY):
        if check_online_healthcare(company_name):
            # auto-learn
            df = pd.read_csv(MERCHANT_FILE)
            df.loc[len(df)] = [company_raw]
            df.to_csv(MERCHANT_FILE, index=False)
            return True

    return False

# ==============================
# PROCESSING
# ==============================

def get_row_value(row, names):
    for name in names:
        if name in row and pd.notna(row[name]):
            return row[name]
    return ""


def process_csv_or_excel(file_path):
    df = pd.read_csv(file_path) if file_path.suffix == ".csv" else pd.read_excel(file_path)

    rows = []

    for _, row in df.iterrows():
        company = clean_company_name(get_row_value(row, ["Description","Merchant","Name"]))
        date = get_row_value(row, ["Date"])
        amount = clean_amount(get_row_value(row, ["Amount"]))

        if company and amount is not None and is_healthcare_company(company):
            rows.append({"Date": date, "Company": company, "Amount": amount})

    return pd.DataFrame(rows)


def parse_pdf_transaction_line(line):
    # Match: 03/06/25 UMA PHARMACY AMHERST, MA 5.68
    pattern = r"(\d{2}/\d{2}/\d{2})\s+(.+?)\s+(-?\d+\.\d{2})\s*$"
    match = re.search(pattern, line)

    if not match:
        return None

    date = match.group(1)
    amount = float(match.group(3))

    # Extract company by removing location (last comma section)
    desc = match.group(2)
    company = desc.split(",")[0]  # remove "AMHERST, MA"

    return {
        "Date": date,
        "Company": clean_company_name(company),
        "Amount": amount
    }


def process_pdf(file_path):
    rows = []
    in_health_section = False

    with pdfplumber.open(file_path) as pdf:
        for page in pdf.pages:
            text = page.extract_text() or ""

            for line in text.split("\n"):

                # START section
                if any(x in line for x in ["Health Care", "Pharmacy", "Health $"]):
                    in_health_section = True
                    continue

                # STOP section
                if any(x in line for x in [
                    "Travel and Transportation",
                    "Services",
                    "Merchandise",
                    "Entertainment",
                    "Food Store"
                ]):
                    in_health_section = False

                parsed = parse_pdf_transaction_line(line)

                if parsed:
                    if in_health_section or is_healthcare_company(parsed["Company"]):
                        rows.append(parsed)

    return pd.DataFrame(rows)

def extract_pdf_totals(file_path):
    health_total = None
    pharmacy_total = None

    with pdfplumber.open(file_path) as pdf:
        for page in pdf.pages:
            text = page.extract_text() or ""

            # Match: Health Care $702.74
            health_match = re.search(r"Health Care\s+\$?([\d,]+\.\d{2})", text)
            if health_match:
                health_total = float(health_match.group(1).replace(",", ""))

            # Match: Pharmacy $163.38
            pharm_match = re.search(r"Pharmacy\s+\$?([\d,]+\.\d{2})", text)
            if pharm_match:
                pharmacy_total = float(pharm_match.group(1).replace(",", ""))

    return health_total, pharmacy_total

def reconcile_totals(df, pdf_file):
    extracted_total = df["Amount"].sum()

    health_total, pharmacy_total = extract_pdf_totals(pdf_file)

    expected_total = 0
    if health_total:
        expected_total += health_total
    if pharmacy_total:
        expected_total += pharmacy_total

    print("\n===== RECONCILIATION =====")
    print(f"Extracted total: ${extracted_total:.2f}")
    print(f"Expected total:  ${expected_total:.2f}")

    diff = round(expected_total - extracted_total, 2)

    if abs(diff) < 0.01:
        print("✅ MATCH — All transactions captured\n")
    else:
        print(f"❌ MISMATCH — Missing ${diff:.2f}\n")

def process_numbers(file_path):
    doc = Document(str(file_path))
    rows = []

    for sheet in doc.sheets:
        for table in sheet.tables:
            data = table.rows(values_only=True)

            for row in data[1:]:
                # Combine entire row into text
                row_text = " ".join([str(v) for v in row if v is not None])

                if not row_text.strip():
                    continue

                first_line = row_text.split("\n")[0].strip()

                # Clean company
                company = clean_company_name(first_line)

                date_match = re.search(r"\d{2}/\d{2}/\d{4}", row_text)
                date = date_match.group(0) if date_match else ""

                nums = [float(v) for v in row if isinstance(v, (int, float)) and abs(v) > 0.01]

                amount = max(nums, key=abs) if nums else 0.0

                if not company:
                    continue

                if is_healthcare_company(company):
                    rows.append({
                        "Date": date,
                        "Company": company,
                        "Amount": amount
                    })

    return pd.DataFrame(rows)

# ==============================
# OUTPUT
# ==============================

def append_to_master(new_data):
    if new_data.empty:
        return

    if MASTER_FILE.exists():
        existing = pd.read_excel(MASTER_FILE)
        combined = pd.concat([existing, new_data], ignore_index=True)
    else:
        combined = new_data

    combined = combined.drop_duplicates(subset=["Date","Company","Amount"])

    total = combined["Amount"].sum()

    total_row = pd.DataFrame([{
        "Date": "",
        "Company": "TOTAL",
        "Amount": round(total, 2)
    }])

    combined = pd.concat([combined, total_row], ignore_index=True)
    combined.to_excel(MASTER_FILE, index=False)

# ==============================
# RUNNER
# ==============================

def process_file(file_path):
    log(f"Processing {file_path.name}")

    suffix = file_path.suffix.lower()

    if suffix in [".csv",".xlsx",".xls"]:
        data = process_csv_or_excel(file_path)
    elif suffix == ".pdf":
        data = process_pdf(file_path)
    elif suffix == ".numbers":
        data = process_numbers(file_path)
    else:
        return

    if suffix == ".pdf":
        data = process_pdf(file_path)

    if not data.empty:
        reconcile_totals(data, file_path)

    append_to_master(data)

    shutil.move(str(file_path), PROCESSED_DIR / file_path.name)


def process_existing_files():
    for f in sorted(INPUT_DIR.iterdir()):
        if f.suffix.lower() in SUPPORTED_FILES:
            process_file(f)


if __name__ == "__main__":
    ensure_folders()
    process_existing_files()