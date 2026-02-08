import re
import os

def clean_number(value_str):
    """Mengambil angka dari string."""
    if not value_str: return 0.0
    text = str(value_str).replace('\n', ' ').strip()
    match = re.search(r'[\d,.]+', text)
    if match:
        clean_num = match.group(0).replace(',', '') # menghapus ribuan
        try:
            return float(clean_num)
        except:
            return 0.0
    return 0.0

def clean_db_cr_flag(value_str):
    """Mencari indikator D atau C."""
    if not value_str: return None
    text = str(value_str).replace('\n', ' ').strip().upper()
    
    if ' D' in text or text.endswith('D') or text == 'D':
        return 'D'
    elif ' C' in text or text.endswith('C') or text == 'C':
        return 'C'
    return None

def extract_bni_data(pdf_path, output_excel):
    import pdfplumber
    import pandas as pd

    data_rows = []
    print(f"\nMemproses file: {pdf_path}...")
    
    try:
        with pdfplumber.open(pdf_path) as pdf:
            for page_num, page in enumerate(pdf.pages):
                table = page.extract_table({
                    "vertical_strategy": "lines", 
                    "horizontal_strategy": "lines",
                    "snap_tolerance": 3,
                })
                
                if not table: continue

                # --- DYNAMIC HEADER MAPPING ---
                header_idx = -1
                col_map = {}
                
                for i, row in enumerate(table):
                    row_str = [str(cell).lower().strip() if cell else '' for cell in row]
                    
                    if "no." in row_str and ("post date" in row_str or "posting date" in row_str):
                        header_idx = i
                        for col_i, text in enumerate(row_str):
                            if "journal" in text: col_map['journal'] = col_i
                            elif text == "no." or text == "no": col_map['no'] = col_i
                            elif "date" in text: col_map['date'] = col_i
                            elif "branch" in text: col_map['branch'] = col_i
                            elif "description" in text: col_map['desc'] = col_i
                            elif "balance" in text: col_map['balance'] = col_i
                            
                            if "amount" in text: col_map['amount'] = col_i
                            if "db/cr" in text: col_map['db_cr'] = col_i
                        
                        # Fallback mappings
                        if 'amount' in col_map and 'db_cr' not in col_map:
                            col_map['db_cr'] = col_map['amount']
                        if 'db_cr' in col_map and 'amount' not in col_map:
                            col_map['amount'] = col_map['db_cr']
                            
                        break
                
                if header_idx == -1: continue
                
                # --- PROSES DATA ---
                current_record = None
                
                for row in table[header_idx+1:]:
                    safe_row = [cell if cell is not None else "" for cell in row]
                    
                    # Cek Kolom No
                    val_no_idx = col_map.get('no', 0)
                    val_no = safe_row[val_no_idx].strip() if val_no_idx < len(safe_row) else ""
                    
                    if val_no:
                        # Simpan record sebelumnya
                        if current_record: data_rows.append(current_record)
                        
                        # --- AMBIL DATA ---
                        val_date = safe_row[col_map.get('date', 1)].split('\n')[0]
                        val_branch = safe_row[col_map.get('branch', 2)].replace('\n', ' ').strip()
                        val_journal = safe_row[col_map.get('journal', 3)].replace('\n', '').strip()
                        
                        # --- LOGIKA AMOUNT & D/C ---
                        idx_amt = col_map.get('amount', 5)
                        idx_dbcr = col_map.get('db_cr', 5)
                        
                        raw_amt_str = safe_row[idx_amt] if idx_amt < len(safe_row) else ""
                        raw_dbcr_str = safe_row[idx_dbcr] if idx_dbcr < len(safe_row) else ""
                        
                        amount_val = clean_number(raw_amt_str)
                        
                        db_cr_flag = clean_db_cr_flag(raw_dbcr_str)
                        if not db_cr_flag and idx_amt != idx_dbcr:
                            db_cr_flag = clean_db_cr_flag(raw_amt_str)

                        # D PDF -> Credit 
                        # C PDF -> Debit 
                        debit = 0.0
                        credit = 0.0
                        
                        if db_cr_flag == 'D':
                            credit = amount_val
                        elif db_cr_flag == 'C':
                            debit = amount_val
                        
                        # Balance
                        raw_balance = safe_row[col_map.get('balance', 6)]
                        balance_val = clean_number(raw_balance)

                        current_record = {
                            "No": val_no,
                            "Posting Date": val_date,
                            "Remark": val_branch,
                            "Reference No": val_journal,
                            "Debit": debit,
                            "Credit": credit,
                            "Balance": balance_val
                        }
                    else:
                        if current_record:
                            idx_branch = col_map.get('branch', 2)
                            if idx_branch < len(safe_row):
                                extra_txt = safe_row[idx_branch].replace('\n', ' ').strip()
                                if extra_txt:
                                    current_record['Remark'] += " " + extra_txt
                
                if current_record:
                    data_rows.append(current_record)

    except Exception as e:
        print(f"Terjadi kesalahan saat membaca PDF: {e}")
        return 0

    if not data_rows:
        print("Tidak ada data yang ditemukan atau format tabel tidak sesuai.")
        return 0

    # Export ke Excel
    try:
        df = pd.DataFrame(data_rows)
        df['Remark'] = df['Remark'].str.replace(r'\s+', ' ', regex=True).str.strip()
        
        print("\nPreview 5 Data Teratas:")
        print(df.head())
        
        df.to_excel(output_excel, index=False)
        print(f"\n[SUKSES] Data berhasil diekstrak dan disimpan di: {output_excel}")
        return len(data_rows)
        
    except Exception as e:
        print(f"Gagal menyimpan file Excel: {e}")
        return 0

# --- BLOCK EKSEKUSI UTAMA ---
if __name__ == "__main__":
    print("=== BNI PDF STATEMENT EXTRACTOR ===")
    print("Pastikan file PDF berada di folder yang sama atau masukkan full path.")
    
    while True:
        pdf_input = input("\nMasukkan nama/path file PDF (contoh: statement.pdf): ").strip()
        pdf_input = pdf_input.replace('"', '').replace("'", "")
        
        if os.path.exists(pdf_input) and pdf_input.lower().endswith('.pdf'):
            break
        else:
            print(f"[ERROR] File '{pdf_input}' tidak ditemukan atau bukan file PDF. Coba lagi.")

    output_input = input("Masukkan nama file output (Tekan Enter untuk default 'hasil_transaksi.xlsx'): ").strip()
    
    if not output_input:
        output_input = "hasil_transaksi.xlsx"
    
    if not output_input.lower().endswith('.xlsx'):
        output_input += '.xlsx'

    extract_bni_data(pdf_input, output_input)