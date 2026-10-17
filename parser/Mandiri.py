import re
from typing import Any, Optional

def parse(pdf_path: str, output_excel_path: str, password: str | None = None, settings: Any = None) -> int:
    """Ekstraksi mutasi Bank Mandiri dari PDF ke file Excel."""
    import pdfplumber
    import pandas as pd

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
        from utils import swap_columns
        swap_columns(df, col_debit, col_credit)
    else:
        print("Peringatan: Kolom Debit atau Credit tidak ditemukan secara otomatis.")

    from exporter import export_dataframe_to_excel
    return export_dataframe_to_excel(df, output_excel_path, settings=settings)

# Backward compatibility alias
process_bank_statement = parse


def parse_ocr(
    pdf_path: str,
    output_excel_path: str,
    password: str | None = None,
    progress_callback: Any = None,
    settings: Any = None
) -> int:
    """Ekstraksi mutasi Mandiri via OCR lokal (menggunakan parser Livin/Mandiri OCR)."""
    from parser.Livin import parse_ocr as livin_parse_ocr
    return livin_parse_ocr(pdf_path, output_excel_path, password=password, progress_callback=progress_callback, settings=settings)



if __name__ == "__main__":
    import sys
    if len(sys.argv) > 2:
        process_bank_statement(sys.argv[1], sys.argv[2])
    else:
        print("Penggunaan: python Mandiri.py <input.pdf> <output.xlsx>")