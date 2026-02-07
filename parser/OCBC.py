import pdfplumber
import pandas as pd
import os
import re

def is_date(string):
    """Cek apakah string berisi pola tanggal (angka/angka)"""
    if not string: return False
    return bool(re.search(r'\d{1,2}/\d{1,2}', str(string)))

def is_new_transaction(row):
    """
    Menentukan apakah baris ini adalah awal dari transaksi baru.
    Kriteria:
    1. Kolom pertama memiliki tanggal.
    2. ATAU baris mengandung kata kunci 'Beginning Balance'/'Saldo Awal'.
    """
    col_date = row[0] if len(row) > 0 else ""
    if is_date(col_date):
        return True

    row_text = " ".join([str(cell).upper() for cell in row])
    keywords = ["BEGINNING BALANCE", "SALDO AWAL", "SALDO SEBELUMNYA", "BROUGHT FORWARD"]
    
    if any(keyword in row_text for keyword in keywords):
        return True
        
    return False

def process_ocbc_final(pdf_path, output_excel_path, password=None):
    print(f"Membaca file: {pdf_path}...")
    
    all_raw_rows = []
    
    # 1. EKSTRAKSI DATA
    with pdfplumber.open(pdf_path, password=password) as pdf:
        for page in pdf.pages:
            table = page.extract_table()
            if table:
                cleaned_table = []
                for row in table:
                    # Bersihkan spasi dan newline
                    cleaned_row = [str(cell).replace('\n', ' ').strip() if cell else "" for cell in row]
                    cleaned_table.append(cleaned_row)
                all_raw_rows.extend(cleaned_table)

    if not all_raw_rows:
        print("Data kosong atau tidak terbaca.")
        return 0

    merged_data = []
    current_transaction = None
    header_found = False
    headers = []

    for row in all_raw_rows:
        if not any(row): continue

        row_str = " ".join(row).upper()
        # Validasi Ketat: OCBC harus punya TRANS, URAIAN, dan VALUTA
        if 'TRANS' in row_str and ('URAIAN' in row_str or 'DESCRIPTION' in row_str) and 'VALUTA' in row_str:
            if not header_found:
                headers = row
                header_found = True
            continue 
        
        # Guard: Jangan proses baris apa pun jika header OCBC belum ditemukan
        if not header_found:
            continue

        if is_new_transaction(row):
            if current_transaction:
                merged_data.append(current_transaction)
            
            current_transaction = row[:]
            
        else:
            if current_transaction:
                desc_part = row[2] if len(row) > 2 else ""

                if desc_part:
                    current_transaction[2] += " " + desc_part
            else:
                pass

    if current_transaction:
        merged_data.append(current_transaction)

    if not header_found or not merged_data:
        headers = ["TGL TRANS", "TGL VALUTA", "URAIAN", "DEBET", "KREDIT", "SALDO"]
    

    if merged_data:
        max_cols = max(len(x) for x in merged_data)
        while len(headers) < max_cols:
            headers.append(f"EXTRA_{len(headers)}")
        headers = headers[:max_cols]

    df = pd.DataFrame(merged_data, columns=headers)
    
    print("Menukar posisi Debit dan Kredit...")
    
    col_debit = next((c for c in df.columns if 'DEB' in c.upper()), None)
    col_credit = next((c for c in df.columns if 'KRE' in c.upper() or 'CRE' in c.upper()), None)

    if col_debit and col_credit:
        temp = df[col_debit].copy()
        df[col_debit] = df[col_credit]
        df[col_credit] = temp
    else:
        print("[Warning] Kolom Debit/Kredit tidak ditemukan otomatis. Penukaran dilewati.")

    try:
        df.to_excel(output_excel_path, index=False)
        print(f"Selesai! File tersimpan: {output_excel_path}")
        return len(merged_data)
    except Exception as e:
        print(f"Gagal menyimpan file Excel: {e}")
        return 0

if __name__ == "__main__":
    raw_input = input("Drag & drop file PDF ke sini: ")
    file_pdf = raw_input.strip().strip('"').strip("'")
    
    if os.path.exists(file_pdf):
        base_name = os.path.basename(file_pdf)
        file_out = f"Hasil_Final_{os.path.splitext(base_name)[0]}.xlsx"
        
        try:
            process_ocbc_final(file_pdf, file_out)
        except Exception as e:
            print(f"Error detail: {e}")
    else:
        print("File tidak ditemukan.")