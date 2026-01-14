import pdfplumber
import pandas as pd
import re

def process_bank_statement(pdf_path, output_excel_path):
    print(f"Membaca file: {pdf_path}...")
    
    all_data = []
    
    with pdfplumber.open(pdf_path) as pdf:
        for i, page in enumerate(pdf.pages):
            table = page.extract_table()
            
            if table:
                if i == 0:
                    headers = table[0]
                    headers = [str(h).replace('\n', ' ').strip() for h in headers]
                    data = table[1:]
                else:

                    data = table
                
                for row in data:
                    if row and any(row): 
                        all_data.append(row)

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

    df.to_excel(output_excel_path, index=False)
    print(f"Selesai! Data tersimpan di: {output_excel_path}")

file_pdf = "/Users/maikerudesu/Downloads/Rekening koran oktober 2025 2/IBIZ_203901081020303_20251001_20251031_1762394240800866720.pdf"
file_output = "mandiri1test.xlsx"

if __name__ == "__main__":
    try:
        process_bank_statement(file_pdf, file_output)
    except Exception as e:
        print(f"Terjadi kesalahan: {e}")