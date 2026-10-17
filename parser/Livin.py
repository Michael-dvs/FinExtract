import os
import re
from datetime import datetime
from typing import Any, Optional

from utils import clean_text, swap_columns

def format_date_excel(date_str):
    """Mengubah '01 Dec 2025' menjadi '01/12/2025' dan membuang jam."""
    if not date_str:
        return ""
    clean = str(date_str).split(',')[0].strip()
    match = re.search(r"(\d{1,2}\s+[A-Za-z]{3}\s+\d{4})", clean)
    if match:
        clean = match.group(1)
    
    try:
        dt_obj = datetime.strptime(clean, "%d %b %Y")
        return dt_obj.strftime("%d/%m/%Y")
    except ValueError:
        return clean

def is_date_start(cell_value):
    """Cek apakah sel ini awal transaksi (Tanggal)."""
    if not cell_value: return False
    return bool(re.match(r"^\d{1,2}\s+[A-Za-z]{3}\s+\d{4}", str(cell_value).strip()))

def is_money(val):
    """Cek apakah string terlihat seperti format uang."""
    if not val: return False
    return bool(re.match(r'^[\d,\.\-]+$', str(val).strip()))

def align_bank_row(row_items):
    """Normalisasi kolom dengan strategi jangkar kanan."""
    clean_items = [x for x in row_items if x and str(x).strip() != ""]
    
    if not clean_items:
        return [""] * 6

    posting_date = clean_items[0]
    balance = "0.00"
    credit = "0.00"
    debit = "0.00"
    
    # Ambil 3 item terakhir sebagai angka
    if len(clean_items) >= 4: 
        balance = clean_items[-1]
        
        if is_money(clean_items[-2]):
            credit = clean_items[-2]
            if len(clean_items) > 2 and is_money(clean_items[-3]):
                debit = clean_items[-3]
                middle_items = clean_items[1:-3]
            else:
                middle_items = clean_items[1:-2]
        else:
            middle_items = clean_items[1:-1]
    else:
        return [posting_date, "PARSE ERROR", "", "0", "0", "0"]

    ref_no = "-"
    remark = ""
    
    if middle_items:
        ref_candidate = middle_items[-1]
        ref_no = ref_candidate
        remark = " ".join(middle_items[:-1])
        
        if not remark and ref_no:
            if len(ref_no) > 25 or ' ' in ref_no:
                remark = ref_no
                ref_no = "-"
    else:
        remark = "" 
        ref_no = "-"

    return [posting_date, remark, ref_no, debit, credit, balance]

def process_raw_rows(raw_rows: list, output_excel_path: str) -> int:
    """Memproses raw_rows hasil ekstraksi (baik dari scraping maupun OCR) menjadi file Excel."""
    import pandas as pd

    if not raw_rows:
        return 0

    # --- MERGE ROWS ---
    merged_data = []
    current_transaction = None

    for row in raw_rows:
        col_first = row[0] if len(row) > 0 else ""
        
        if is_date_start(col_first):
            if current_transaction:
                merged_data.append(current_transaction)
            current_transaction = row
        else:
            if current_transaction:
                for idx in range(1, len(row)):
                    if idx < len(current_transaction):
                        text_fragment = row[idx]
                        if text_fragment:
                            current_transaction[idx] += " " + text_fragment

    if current_transaction:
        merged_data.append(current_transaction)

    # --- ALIGN COLUMNS ---
    final_rows = []
    for row in merged_data:
        aligned_row = align_bank_row(row)
        final_rows.append(aligned_row)

    # --- BUILD DATAFRAME ---
    fixed_headers = ["Posting Date", "Remark", "Reference No", "Debit", "Credit", "Balance"]
    df = pd.DataFrame(final_rows, columns=fixed_headers)

    if "Posting Date" in df.columns:
        df["Posting Date"] = df["Posting Date"].apply(format_date_excel)

    # Menukar isi kolom Debit dengan Credit, dan sebaliknya
    print("Menukar posisi kolom Debit dan Credit...")
    swap_columns(df, "Debit", "Credit")

    # Simpan
    try:
        df.to_excel(output_excel_path, index=False)
        return len(final_rows)
    except PermissionError:
        print(f"ERROR: Gagal menyimpan file ke '{output_excel_path}'.")
        print("SOLUSI: Pastikan file Excel tersebut TIDAK SEDANG DIBUKA. Tutup file Excel lalu coba lagi.")
        return 0
    except Exception as e:
        print(f"Gagal menyimpan file Excel: {e}")
        return 0


def parse(pdf_path: str, output_excel_path: str | None = None, password: str | None = None) -> int:
    """Ekstraksi mutasi Livin' by Mandiri dari PDF ke file Excel via scraping."""
    import pdfplumber

    try:
        if not output_excel_path:
            base_name = os.path.splitext(pdf_path)[0]
            output_excel_path = f"{base_name}_livin.xlsx"
        elif os.path.isdir(output_excel_path):
            base_name = os.path.splitext(os.path.basename(pdf_path))[0]
            output_excel_path = os.path.join(output_excel_path, f"{base_name}_livin.xlsx")

        print(f"[LIVIN] Memproses (Scraping): {os.path.basename(pdf_path)}...")
        
        raw_rows = []
        
        table_settings = {
            "vertical_strategy": "text", 
            "horizontal_strategy": "text",
            "text_x_tolerance": 5, 
            "snap_tolerance": 4,
        }

        with pdfplumber.open(pdf_path) as pdf:
            for i, page in enumerate(pdf.pages):
                table = page.extract_table(table_settings)
                
                if table:
                    start_index = 0
                    for idx, row in enumerate(table):
                        row_str = " ".join([str(c).lower() for c in row if c])
                        if "posting date" in row_str or "tanggal" in row_str:
                            start_index = idx + 1
                            break
                    
                    data_page = table[start_index:]
                    for row in data_page:
                        cleaned_row = [clean_text(c) for c in row]
                        if any(cleaned_row):
                            raw_rows.append(cleaned_row)

        if not raw_rows:
            return 0

        return process_raw_rows(raw_rows, output_excel_path)

    except Exception as e:
        import traceback
        traceback.print_exc()
        raise


def parse_ocr(
    pdf_path: str,
    output_excel_path: str | None = None,
    password: str | None = None,
    progress_callback: Any = None
) -> int:
    """Ekstraksi mutasi Livin' by Mandiri dari PDF ke file Excel via OCR lokal."""
    try:
        from ocr_engine import extract_pdf_ocr_raw_rows, LIVIN_COLUMNS_FRACTION

        if not output_excel_path:
            base_name = os.path.splitext(pdf_path)[0]
            output_excel_path = f"{base_name}_livin.xlsx"
        elif os.path.isdir(output_excel_path):
            base_name = os.path.splitext(os.path.basename(pdf_path))[0]
            output_excel_path = os.path.join(output_excel_path, f"{base_name}_livin.xlsx")

        print(f"[LIVIN] Memproses (OCR): {os.path.basename(pdf_path)}...")

        raw_rows = extract_pdf_ocr_raw_rows(
            pdf_path=pdf_path,
            columns_fraction=LIVIN_COLUMNS_FRACTION,
            dpi=200,
            progress_callback=progress_callback
        )

        cleaned_raw_rows = []
        for r in raw_rows:
            cleaned_row = [clean_text(c) for c in r]
            if any(cleaned_row):
                cleaned_raw_rows.append(cleaned_row)

        if not cleaned_raw_rows:
            return 0

        return process_raw_rows(cleaned_raw_rows, output_excel_path)

    except Exception as e:
        import traceback
        traceback.print_exc()
        return 0


# Backward compatibility alias
process_livin_statement = parse