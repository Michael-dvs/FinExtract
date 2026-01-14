import pdfplumber
import pandas as pd
import re
import os

def extract_bri_text(pdf_path, excel_path):
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
    
    # Error handling permission
    try:
        df.to_excel(excel_path, index=False)
        print(f"Data berhasil diekspor ke {excel_path} dengan {len(df)} baris.")
    except PermissionError:
        print(f"ERROR: Gagal menyimpan file ke '{excel_path}'.")
        print("SOLUSI: Pastikan file Excel tersebut TIDAK SEDANG DIBUKA. Tutup file Excel lalu coba lagi.")
    except Exception as e:
        print(f"Terjadi error lain saat menyimpan: {e}")

if __name__ == "__main__":

    #for windows:
    #pdf_path = r"C:\Users\Maikeru\Downloads\Rekap Rekening koran jan-june 2025\Rekap Rekening koran jan-june 2025\Rekening koran juli 2025\BRI\Mutasi Juli Rekening BRI 203901000446308.pdf"
    #excel_path = input("Masukkan nama file Excel output (misal: hasil_bri.xlsx): ")+".xlsx"
    #extract_bri_text(pdf_path, excel_path)

    #for macOS/Linux:
    folder_path = "/Users/maikerudesu/Downloads/Rekening koran oktober 2025"

    if os.path.exists(folder_path):
        print(f"Membaca direktori: {folder_path}")
        for filename in os.listdir(folder_path):
            if filename.lower().endswith(".pdf"):
                full_pdf_path = os.path.join(folder_path, filename)
                output_excel = os.path.join(folder_path, os.path.splitext(filename)[0] + ".xlsx")
                print(f"Memproses: {filename}")
                extract_bri_text(full_pdf_path, output_excel)
    else:
        print(f"Folder tidak ditemukan: {folder_path}")