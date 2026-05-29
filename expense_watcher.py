import re
from pathlib import Path
import pdfplumber
import pandas as pd
import sys
import os

# ==============================
# BASE PATH (WORKS FOR EXE + SCRIPT)
# ==============================
if getattr(sys, 'frozen', False):
    BASE_DIR = Path(sys.executable).parent
else:
    BASE_DIR = Path(__file__).resolve().parent

INPUT_DIR = BASE_DIR / "input_files"
OUTPUT_DIR = BASE_DIR / "output"
OUTPUT_FILE = OUTPUT_DIR / "Healthcare_Expenses_Master.xlsx"

# ==============================
# 🚨 ENSURE DIRECTORIES (BULLETPROOF)
# ==============================
def ensure_directories():
    try:
        INPUT_DIR.mkdir(parents=True, exist_ok=True)
        OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

        print("\n📂 DIRECTORY CHECK")
        print(f"BASE DIR: {BASE_DIR.resolve()}")
        print(f"INPUT DIR: {INPUT_DIR.resolve()}")
        print(f"OUTPUT DIR: {OUTPUT_DIR.resolve()}")

        # sanity check
        print("\n📁 Contents of BASE DIR:")
        print(list(BASE_DIR.iterdir()))

    except Exception as e:
        print(f"❌ Failed to create directories: {e}")

# ==============================
# CLEAN COMPANY NAME
# ==============================
def clean_company_name(text):
    text = text.upper()

    text = re.sub(r"^(CHECKCARD|POS|ACH|DEBIT)\s+\d+\s*", "", text)
    text = re.sub(r"\d{6,}", "", text)
    text = re.sub(r"\d{3}-\d{3}-\d{4}", "", text)
    # Fix merged pharmacy names
    text = re.sub(r"(CVS)(PHARMACY)", r"\1 \2", text)
    text = re.sub(r"(WALGREENS)(PHARMACY)", r"\1 \2", text)

    # fix duplicated words
    text = re.sub(r"\b(\w+)\s+\1\b", r"\1", text)

    text = re.sub(r"\s+", " ", text)

    return text.strip()

# ==============================
# FILTER OUT NON-REAL TRANSACTIONS
# ==============================
def is_valid_company(company):
    bad = [
        "ATM", "WITHDRWL", "ZELLE", "TRANSFER",
        "DEPOSIT", "VENMO", "PAYMENT",
        "BANK OF AMERICA", "UMASS STORE"
    ]
    return not any(b in company for b in bad)

def is_false_positive(company):
    bad = [
        "MASS BAY",      # bookstore
        "BKST",          # bookstore shorthand
        "TAX",
        "DEPARTME",
        "REFUND",
        "TREASURY"
    ]
    return any(b in company for b in bad)

# ==============================
# HEALTHCARE FILTER (BALANCED)
# ==============================
def is_healthcare(company):
    text = company.upper()

    keywords = [
        # pharmacies
        "CVS", "WALGREEN", "RITE AID", "PHARM",

        # providers (specific only)
        "ATRIUS",
        "MASS GENERAL",
        "BRIGHAM",
        "PARTNERS",
        "NEWTON WELLESLEY",
        "MOUNT AUBURN",
        "UMASS HEALTH",
        "UMA PHARMACY",

        # specialties
        "DERM",
        "VISION",
        "DENTAL",
        "CARE",
        "CLINIC",
        "HOSPITAL",

        # labs
        "LABCORP", "QUEST",

        # retail medical
        "WARBY"
    ]

    return any(k in text for k in keywords)

# ==============================
# ROBUST PDF PARSER
# ==============================
def parse_pdf(file_path):
    rows = []

    try:
        with pdfplumber.open(file_path) as pdf:
            for page in pdf.pages:
                text = page.extract_text()

                if not text:
                    continue

                for line in text.split("\n"):
                    line = line.strip()

                    # =========================
                    # PRIMARY REGEX PARSE
                    # =========================
                    match = re.search(
                        r"(\d{2}/\d{2}/\d{2})\s+(.*?)\s+(-?\d+\.\d{2})\s*$",
                        line
                    )

                    if match:
                        date = match.group(1)
                        company_raw = match.group(2)
                        amount_raw = match.group(3)

                        # Remove CR/DR or any non-numeric characters
                        amount_clean = re.sub(r"[^\d\.\-]", "", amount_raw)

                        try:
                            amount = float(amount_clean)
                        except:
                            continue

                        company = clean_company_name(company_raw)

                        rows.append({
                            "Date": date,
                            "Company": company,
                            "Amount": amount
                        })

                    # =========================
                    # FALLBACK PARSE
                    # =========================
                    else:
                        parts = line.split()

                        if len(parts) >= 3:
                            if (
                                re.match(r"\d{2}/\d{2}/\d{2}", parts[0]) and
                                re.match(r"-?\d+\.\d{2}", parts[-1])
                            ):
                                date = parts[0]
                                amount_raw = parts[-1]
                                amount_clean = re.sub(r"[^\d\.\-]", "", amount_raw)

                                try:
                                    amount = float(amount_clean)
                                except:
                                    continue

                                company_raw = " ".join(parts[1:-1])

                                company = clean_company_name(company_raw)

                                rows.append({
                                    "Date": date,
                                    "Company": company,
                                    "Amount": amount
                                })

    except Exception as e:
        print(f"❌ Error parsing {file_path.name}: {e}")

    return pd.DataFrame(rows)

# ==============================
# MAIN PROCESSOR
# ==============================
def process_all_files():
    print("\n🚀 Processing files...")
    print(f"Looking in: {INPUT_DIR.resolve()}")

    files = list(INPUT_DIR.glob("*.pdf"))

    print(f"\n📄 FILES FOUND: {len(files)}")

    if not files:
        print("❌ No PDF files found in input_files/")
        return

    all_data = []

    for file in sorted(files):
        print(f"\n--- Processing {file.name} ---")

        df = parse_pdf(file)

        if df.empty:
            print("⚠️ No transactions found")
            continue

        print(f"✅ Parsed {len(df)} transactions")

        df["Source File"] = file.name
        all_data.append(df)

    if not all_data:
        print("❌ No data extracted")
        return

    final_df = pd.concat(all_data, ignore_index=True)

    print(f"\n📊 Total transactions before filtering: {len(final_df)}")

    # ==============================
    # CLEAN DATA
    # ==============================
    final_df = final_df[
        final_df["Company"].apply(is_valid_company) &
        ~final_df["Company"].apply(is_false_positive)
    ]

    print(f"After removing junk: {len(final_df)}")

    # ==============================
    # HEALTHCARE FILTER
    # ==============================
    final_df["Is Healthcare"] = final_df["Company"].apply(is_healthcare)

    healthcare_df = final_df[final_df["Is Healthcare"]].copy()

    print(f"Healthcare transactions found: {len(healthcare_df)}")

    if healthcare_df.empty:
        print("⚠️ No healthcare transactions found")
        return

    healthcare_df = healthcare_df.drop(columns=["Is Healthcare"])

    # ==============================
    # SAVE OUTPUT
    # ==============================
    try:
        healthcare_df.to_excel(OUTPUT_FILE, index=False)
        print(f"\n✅ OUTPUT SAVED → {OUTPUT_FILE.resolve()}")
        print(f"💰 Total Healthcare Spend: ${round(healthcare_df['Amount'].sum(), 2)}")

    except Exception as e:
        print(f"❌ Failed to save Excel: {e}")

# ==============================
# RUN
# ==============================
if __name__ == "__main__":
    print("\n=== Healthcare Expense Tool Starting ===")
    ensure_directories()
    process_all_files()
    print("\n=== Done ===")