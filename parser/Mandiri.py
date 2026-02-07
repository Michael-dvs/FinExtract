import pdfplumber
import pandas as pd
import re

def process_bank_statement(pdf_path, output_excel_path):
    print(f"Membaca file: {pdf_path}...")
    
    all_data = []
    headers = None
    
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            table = page.extract_table()
            
            if table:
                if headers is None:
                    temp_headers = [str(h).replace('\n', ' ').strip() for h in table[0]]
                    header_str = " ".join(temp_headers).upper()
                    
                    # Validasi Signature: 
                    # 1. Format Rekening Koran: Ada (KETERANGAN/DESCRIPTION) DAN (CABANG/BRANCH)
                    # 2. Format Account Statement: Ada (POSTING DATE) DAN (REMARK)
                    if (("KETERANGAN" in header_str or "DESCRIPTION" in header_str) and ("CABANG" in header_str or "BRANCH" in header_str)) or \
                       ("POSTING DATE" in header_str and "REMARK" in header_str):
                        headers = temp_headers
                        data = table[1:]
                    else:
                        continue
                else:
                    data = table
                
                for row in data:
                    if row and any(row): 
                        all_data.append(row)

    if headers is None or not all_data:
        print("Gagal menemukan tabel atau data Mandiri yang valid.")
        return 0

    df = pd.DataFrame(all_data, columns=headers)
    
    col_debit = next((c for c in df.columns if 'Debit' in c), None)
    col_credit = next((c for c in df.columns if 'Credit' in c), None)

    if col_debit and col_credit:
        print(f"Menukar data kolom '{col_debit}' dengan '{col_credit}'...")

        temp_debit_data = df[col_debit].copy()
        
        df[col_debit] = df[col_credit] 
        df[col_credit] = temp_debit_data 
    else:
        print("Peringatan: Kolom Debit atau Credit tidak ditemukan secara otomatis.")

    try:
        df.to_excel(output_excel_path, index=False)
        print(f"Selesai! Data tersimpan di: {output_excel_path}")
        return len(all_data)
    except Exception as e:
        print(f"Gagal menyimpan file Excel: {e}")
        return 0

file_pdf = "/Users/maikerudesu/Downloads/Rekening koran oktober 2025 2/IBIZ_203901081020303_20251001_20251031_1762394240800866720.pdf"
file_output = "mandiri1test.xlsx"

if __name__ == "__main__":
    try:
        process_bank_statement(file_pdf, file_output)
    except Exception as e:
        print(f"Terjadi kesalahan: {e}")