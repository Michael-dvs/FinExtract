import os
import re
from datetime import datetime

def clean_text(text):
    """Membersihkan teks dari newline dan spasi berlebih."""
    if text is None:
        return ""
    text_str = str(text).replace('\n', ' ')
    return re.sub(r'\s+', ' ', text_str).strip()

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

def process_livin_statement(pdf_path):
    import pdfplumber
    import pandas as pd

    try:
        base_name = os.path.splitext(pdf_path)[0]
        output_excel_path = f"{base_name}_livin.xlsx"
        print(f"[LIVIN] Memproses: {os.path.basename(pdf_path)}...")
        
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
            return False, "Gagal mengekstrak data tabel."

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
        temp_debit = df["Debit"].copy()
        df["Debit"] = df["Credit"]
        df["Credit"] = temp_debit

        # Simpan
        df.to_excel(output_excel_path, index=False)
        return len(final_rows)

    except Exception as e:
        import traceback
        traceback.print_exc()
        return 0