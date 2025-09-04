import pdfplumber
import pandas as pd
import re

# ---- CONFIG ----
HEADER_NAMES = [
    "Dates", "Verified", "Code", "D", "A", "LK", "ML", "MID", "JK",
    "Fish Ko", "Movie", "Nomenclature", "DSI", "LK", "BN", "OP"
]

# ------------------- PDF Extract -------------------

def extract_pdf_table(pdf_file):
    """Extract tabular rows from PDF into list of dicts"""
    rows = []
    with pdfplumber.open(pdf_file) as pdf:
        for page in pdf.pages:
            words = page.extract_words(
                keep_blank_chars=False,
                use_text_flow=True,
                x_tolerance=2,
                y_tolerance=3,
            )
            text_lines = page.extract_text().split("\n") if page.extract_text() else []
            for ln in text_lines:
                ln = ln.strip()
                if not ln or re.fullmatch(r"[-\s]+", ln):  # skip dashed lines
                    continue
                parts = ln.split()
                # crude matching into columns (you can refine if needed)
                row = {}
                for i, col in enumerate(HEADER_NAMES):
                    if i < len(parts):
                        row[col] = parts[i]
                    else:
                        row[col] = ""
                rows.append(row)
    return rows

# ------------------- Comparison -------------------

def compare_pdf_excel(pdf_file, excel_file):
    pdf_rows = extract_pdf_table(pdf_file)
    excel_df = pd.read_excel(excel_file)

    # Ensure Excel has the same headers
    for col in HEADER_NAMES:
        if col not in excel_df.columns:
            excel_df[col] = ""

    excel_rows = excel_df.to_dict(orient="records")

    mismatches = []
    for i, (p, e) in enumerate(zip(pdf_rows, excel_rows)):
        for col in HEADER_NAMES:
            pdf_val = str(p.get(col, "")).strip()
            excel_val = str(e.get(col, "")).strip()
            if pdf_val != excel_val:
                mismatches.append((i+1, col, pdf_val, excel_val))

    if mismatches:
        print("⚠️ Discrepancies found:")
        for row, col, pdf_val, excel_val in mismatches:
            print(f"Row {row}, Column {col}: PDF='{pdf_val}' | Excel='{excel_val}'")
    else:
        print("✅ Perfect match: PDF and Excel tabular data are identical.")

# ------------------- Run -------------------

if __name__ == "__main__":
    pdf_path = input("Enter PDF file path: ").strip()
    excel_path = input("Enter Excel file path: ").strip()
    compare_pdf_excel(pdf_path, excel_path)
