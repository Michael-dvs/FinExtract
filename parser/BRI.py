import re
import os

from typing import Any, Optional

def parse(pdf_path: str, output_excel_path: str, password: str | None = None, settings: Any = None) -> int:
    """Ekstraksi mutasi BRI dari PDF ke file Excel."""
    import pdfplumber
    import pandas as pd
    columns = ["transaction_date", "description", "user_id", "debit", "credit", "balance"]
    data = []

    # Pattern regex
    pattern = re.compile(
        r"^(\d{2}/\d{2}/\d{2})\s+\d{2}:\d{2}:\d{2}\s+(.+?)\s+(\d{7,})?\s*([\d,]+\.\d{2})\s+([\d,]+\.\d{2})\s+([\d,]+\.\d{2})$"
    )
    desc_cont_pattern = re.compile(r"^(?!\d{2}/\d{2}/\d{2}\s+\d{2}:\d{2}:\d{2}).+")

    # Daftar pattern footer
    footer_patterns = [
        re.compile(r"halaman\s+\d+", re.IGNORECASE),
        re.compile(r"saldo akhir", re.IGNORECASE),
        re.compile(r"jumlah\s+mutasi", re.IGNORECASE),
        re.compile(r"rekening\s+koran", re.IGNORECASE),
        re.compile(r"^$"),
        re.compile(r"Created By IBBIZ"),
        re.compile(r"\d{2}/\d{2}/\d{4}\s+\d{2}:\d{2}:\d{2}"),
        re.compile(r"saldo awal", re.IGNORECASE),
        re.compile(r"opening balance", re.IGNORECASE),
        re.compile(r"closing balance", re.IGNORECASE),
        re.compile(r"total transaksi debet", re.IGNORECASE),
        re.compile(r"total debit transaction", re.IGNORECASE),
        re.compile(r"total transaksi kredit", re.IGNORECASE),
        re.compile(r"total credit transaction", re.IGNORECASE),
        re.compile(r"terbilang", re.IGNORECASE),
        re.compile(r"in words", re.IGNORECASE),
        re.compile(r"biaya materai", re.IGNORECASE),
        re.compile(r"revenue stamp paid", re.IGNORECASE),
    ]

    def is_footer(line):
        return any(p.search(line) for p in footer_patterns)

    temp_row = None
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            text = page.extract_text()
            if not text:
                continue

            for line in text.split("\n"):
                line = line.strip()
                if is_footer(line):
                    continue

                match = pattern.match(line)
                if match:
                    if temp_row:
                        data.append(temp_row)
                    
                    groups = match.groups()
                    
                    tgl = groups[0]
                    desc = groups[1]
                    user_id = groups[2]
                    
                    # Mengambil angka mentah dari PDF
                    angka_posisi_kiri = groups[3]
                    angka_posisi_kanan = groups[4]
                    balance = groups[5]
                    
                    debit = angka_posisi_kanan 
                    credit = angka_posisi_kiri 
                                        
                    temp_row = [tgl, desc.strip(), user_id or "", debit, credit, balance]

                elif temp_row and desc_cont_pattern.match(line):
                    temp_row[1] += " " + line.strip()
                    
            if temp_row:
                data.append(temp_row)
                temp_row = None

    df = pd.DataFrame(data, columns=columns)
    
    from exporter import export_dataframe_to_excel
    return export_dataframe_to_excel(df, output_excel_path, settings=settings)


# Backward compatibility alias
extract_bri_text = parse

if __name__ == "__main__":
    import sys
    if len(sys.argv) > 2:
        extract_bri_text(sys.argv[1], sys.argv[2])
    else:
        print("Penggunaan: python BRI.py <input.pdf> <output.xlsx>")