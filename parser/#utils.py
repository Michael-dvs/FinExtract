import pandas as pd
import os
import json

def load_bank_config(bank_name):
    """Membaca konfigurasi spesifik bank dari settings.json"""
    user_home = os.path.expanduser("~")
    config_path = os.path.join(user_home, "FinExtract_Settings.json")
    
    default_config = {}
    
    if os.path.exists(config_path):
        try:
            with open(config_path, 'r') as f:
                data = json.load(f)
                # Ambil config khusus bank tersebut, misal: data["bank_configs"]["Mandiri"]
                return data.get("bank_configs", {}).get(bank_name, {})
        except:
            pass
    return default_config

def save_styled_excel(df, output_path, bank_name="General"):
    """
    df: DataFrame Pandas (Internal Columns)
    output_path: Path output
    bank_name: String (Misal: 'Mandiri', 'BCA', 'BNI') untuk kunci config
    """
    
    # 1. Load Config User untuk Bank ini
    bank_config = load_bank_config(bank_name)
    
    # 2. Filter Kolom (Hide/Show) & Rename
    # Kita hanya akan memasukkan kolom yang (1) Ada di DF DAN (2) Visible=True (atau belum disetting)
    
    rename_map = {}
    final_cols_order = []
    
    # Urutan kolom default sesuai DataFrame
    for col in df.columns:
        col_conf = bank_config.get(col, {})
        
        # Cek Visibility (Default True jika tidak ada setting)
        is_visible = col_conf.get("visible", True)
        
        if is_visible:
            user_label = col_conf.get("label", col.upper())
            rename_map[col] = user_label
            final_cols_order.append(col)
            
    # Buat DF baru hanya dengan kolom yang dipilih user
    df_export = df[final_cols_order].rename(columns=rename_map)

    if df_export.empty:
        print("[Warning] Tidak ada kolom yang dipilih untuk diexport.")
        return

    # 3. Write Excel
    try:
        with pd.ExcelWriter(output_path, engine='xlsxwriter') as writer:
            df_export.to_excel(writer, sheet_name='Data', index=False, startrow=1)
            
            workbook = writer.book
            worksheet = writer.sheets['Data']
            
            # Format Dasar
            header_base_fmt = {'bold': True, 'text_wrap': True, 'valign': 'vcenter', 'border': 1}
            data_base_fmt = {'border': 1, 'valign': 'vcenter'}

            for idx, internal_col in enumerate(final_cols_order):
                conf = bank_config.get(internal_col, {})
                
                # --- HEADER STYLE ---
                bg_color = conf.get("bg_color", "#D7E4BC") 
                font_color = conf.get("font_color", "#000000")
                align = conf.get("align", "center")
                
                header_fmt = workbook.add_format({
                    **header_base_fmt,
                    'bg_color': bg_color,
                    'font_color': font_color,
                    'align': align
                })
                
                worksheet.write(0, idx, rename_map[internal_col], header_fmt)
                
                # --- DATA STYLE ---
                data_fmt_dict = data_base_fmt.copy()
                data_fmt_dict['align'] = align 
                
                col_format = conf.get("format", "text")
                if col_format == "accounting":
                    data_fmt_dict['num_format'] = '_-"Rp"* #,##0.00_-;-"Rp"* #,##0.00_-;_-"Rp"* "-"??_-;_-@_-'
                elif col_format == "number":
                    data_fmt_dict['num_format'] = '#,##0'
                
                data_fmt = workbook.add_format(data_fmt_dict)
                
                width = conf.get("width", 20)
                worksheet.set_column(idx, idx, width, data_fmt)
                
        print(f"[{bank_name}] Excel saved: {output_path}")
        
    except Exception as e:
        print(f"Gagal styling excel: {e}")
        df.to_excel(output_path, index=False)