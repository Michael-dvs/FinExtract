# --- KONFIGURASI TEMA MENYELURUH ---
THEME_CONFIG = {
    "Default (Blue)": {
        "window_bg": ("#FAFAFA", "#242424"),         
        "sidebar_bg": ("#EBEBEB", "#2B2B2B"),        
        "frame_bg": ("#FFFFFF", "#333333"),          
        "text": ("black", "white"),                  
        "btn_main": ("#3B8ED0", "#1F6AA5"),          
        "btn_hover": ("#36719F", "#144870"),         
        "border": ("#979DA2", "#565B5E"),            
        "entry_bg": ("#F9F9FA", "#343638"),          
        "log_info": ("gray30", "gray70")             
    },
    "Special Edition (Pink)": {
        "window_bg": ("#FFF0F5", "#2D1B22"),         
        "sidebar_bg": ("#F8BBD0", "#4A1A2C"),        
        "frame_bg": ("#FCE4EC", "#3D242E"),          
        "text": ("#880E4F", "#F8BBD0"),              
        "btn_main": ("#EC407A", "#D81B60"),          
        "btn_hover": ("#D81B60", "#AD1457"),         
        "border": ("#F48FB1", "#880E4F"),            
        "entry_bg": ("#FFF8FA", "#28181D"),          
        "log_info": ("#880E4F", "#F8BBD0")           
    },
}

AUTO_BANKS = [
    ('parser.BNI', 'extract_bni_data', ['BANK NEGARA INDONESIA', 'BNIDIRECT']),
    ('parser.Mandiri', 'process_bank_statement', ['BANK MANDIRI', 'MANDIRI', 'ACCOUNT STATEMENT']),
    ('parser.Livin', 'process_livin_statement', ['LIVIN BY MANDIRI']),
    ('parser.OCBC', 'process_ocbc_final', ['OCBC NISP', 'BANK OCBC', 'OCBC']),
    ('parser.BRI', 'extract_bri_text', ['BANK RAKYAT INDONESIA', 'BRIDIRECT', 'IBBIZ', 'IBIZ', 'BRI'])
]