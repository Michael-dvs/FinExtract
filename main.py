import os
import sys
import threading
import importlib
import traceback
import subprocess
import tempfile
import tkinter as tk
import queue
import json
from tkinter import filedialog, messagebox

try:
    import customtkinter as ctk
except ImportError:
    raise RuntimeError("customtkinter tidak terpasang. Jalankan: pip install customtkinter")

if sys.platform == "darwin":
    try:
        _original_ctk_font_init = ctk.CTkFont.__init__
        def _new_ctk_font_init(self, family=None, *args, **kwargs):
            if family is None: family = "SF Pro Display"
            _original_ctk_font_init(self, family, *args, **kwargs)
        ctk.CTkFont.__init__ = _new_ctk_font_init
    except Exception: pass

try:
    from pypdf import PdfReader, PdfWriter
except ImportError:
    PdfReader = None
    PdfWriter = None

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

class SettingsDialog(ctk.CTkToplevel):
    """Jendela Popup Pengaturan (Safe Mode - No Grab Set)"""
    def __init__(self, parent):
        super().__init__(parent)
        self.parent = parent
        self.title("Settings")
        self.geometry("400x450")
        self.resizable(False, False)
        self.attributes("-topmost", True) 
        self.focus_force()

        self.current_theme = THEME_CONFIG.get(parent.current_theme_key)
        self.configure(fg_color=self.current_theme["window_bg"])

        self.var_mode = tk.StringVar(value=ctk.get_appearance_mode())
        self.var_theme = tk.StringVar(value=parent.current_theme_key)

        self.lbl_title = ctk.CTkLabel(self, text="Pengaturan", font=ctk.CTkFont(size=20, weight="bold"), 
                                      text_color=self.current_theme["text"])
        self.lbl_title.pack(pady=(20, 10))

        # Frame Mode
        self.frame_mode = ctk.CTkFrame(self, fg_color=self.current_theme["frame_bg"], border_width=1, border_color=self.current_theme["border"])
        self.frame_mode.pack(padx=20, pady=10, fill="x")
        self.lbl_mode = ctk.CTkLabel(self.frame_mode, text="Mode Tampilan:", font=ctk.CTkFont(weight="bold"), text_color=self.current_theme["text"])
        self.lbl_mode.pack(padx=10, pady=(10, 5), anchor="w")
        
        for val in ["System", "Light", "Dark"]:
            rb = ctk.CTkRadioButton(self.frame_mode, text=val, variable=self.var_mode, value=val, command=self.on_mode_change, 
                                    fg_color=self.current_theme["btn_main"], hover_color=self.current_theme["btn_hover"], text_color=self.current_theme["text"])
            rb.pack(padx=20, pady=5, anchor="w")

        # Frame Tema
        self.frame_theme = ctk.CTkFrame(self, fg_color=self.current_theme["frame_bg"], border_width=1, border_color=self.current_theme["border"])
        self.frame_theme.pack(padx=20, pady=10, fill="x")
        self.lbl_theme = ctk.CTkLabel(self.frame_theme, text="Tema Warna:", font=ctk.CTkFont(weight="bold"), text_color=self.current_theme["text"])
        self.lbl_theme.pack(padx=10, pady=(10, 5), anchor="w")

        for val in ["Default (Blue)", "Special Edition (Pink)"]:
            rb = ctk.CTkRadioButton(self.frame_theme, text=val, variable=self.var_theme, value=val, command=self.on_theme_change,
                                    fg_color=self.current_theme["btn_main"], hover_color=self.current_theme["btn_hover"], text_color=self.current_theme["text"])
            rb.pack(padx=20, pady=5, anchor="w")

        self.btn_close = ctk.CTkButton(self, text="Tutup", command=self.destroy, fg_color=self.current_theme["btn_main"], hover_color=self.current_theme["btn_hover"])
        self.btn_close.pack(pady=20)

    def on_mode_change(self):
        self.after(10, lambda: self.parent.change_appearance_mode_event(self.var_mode.get()))

    def on_theme_change(self):
        self.after(10, lambda: self.parent.change_theme_event(self.var_theme.get()))

    def update_colors(self, theme_config):
        try:
            self.current_theme = theme_config
            self.configure(fg_color=theme_config["window_bg"])
            self.lbl_title.configure(text_color=theme_config["text"])
            self.frame_mode.configure(fg_color=theme_config["frame_bg"], border_color=theme_config["border"])
            self.frame_theme.configure(fg_color=theme_config["frame_bg"], border_color=theme_config["border"])
            self.lbl_mode.configure(text_color=theme_config["text"])
            self.lbl_theme.configure(text_color=theme_config["text"])
            self.btn_close.configure(fg_color=theme_config["btn_main"], hover_color=theme_config["btn_hover"])
            
            # Update all children widgets simply
            for widget in self.frame_mode.winfo_children():
                if isinstance(widget, ctk.CTkRadioButton):
                    widget.configure(fg_color=theme_config["btn_main"], hover_color=theme_config["btn_hover"], text_color=theme_config["text"])
            for widget in self.frame_theme.winfo_children():
                if isinstance(widget, ctk.CTkRadioButton):
                    widget.configure(fg_color=theme_config["btn_main"], hover_color=theme_config["btn_hover"], text_color=theme_config["text"])
                    
        except Exception: pass

class PasswordDialog(ctk.CTkToplevel):
    def __init__(self, parent, title="Password", text="Enter password:", theme_data=None):
        super().__init__(parent)
        self.title(title)
        self.transient(parent)
        self.grab_set() 
        self.resizable(False, False)
        self.theme = theme_data if theme_data else THEME_CONFIG["Default (Blue)"]
        self.configure(fg_color=self.theme["window_bg"])

        self.label = ctk.CTkLabel(self, text=text, wraplength=300, justify="left", text_color=self.theme["text"])
        self.label.pack(padx=20, pady=(20, 10), fill="x")

        entry_frame = ctk.CTkFrame(self, fg_color="transparent")
        entry_frame.pack(padx=20, pady=10, fill="x", expand=True)
        self.entry = ctk.CTkEntry(entry_frame, show="*", border_color=self.theme["border"], fg_color=self.theme["entry_bg"], text_color=self.theme["text"])
        self.entry.grid(row=0, column=0, sticky="ew", padx=(0, 5))
        self.entry.bind("<Return>", self._ok_event)
        self.show_hide_button = ctk.CTkButton(entry_frame, text="Show", width=60, command=self._toggle_password, fg_color=self.theme["btn_main"], hover_color=self.theme["btn_hover"])
        self.show_hide_button.grid(row=0, column=1, sticky="e")

        button_frame = ctk.CTkFrame(self, fg_color="transparent")
        button_frame.pack(padx=20, pady=(10, 20), fill="x")
        self.ok_button = ctk.CTkButton(button_frame, text="OK", command=self._ok_event, fg_color=self.theme["btn_main"], hover_color=self.theme["btn_hover"])
        self.ok_button.pack(side="left", fill="x", expand=True, padx=(0,5))
        self.cancel_button = ctk.CTkButton(button_frame, text="Cancel", fg_color="gray", hover_color="gray50", command=self._cancel_event)
        self.cancel_button.pack(side="right", fill="x", expand=True, padx=(5,0))
        self.entry.focus()

    def _toggle_password(self):
        if self.entry.cget("show") == "*":
            self.entry.configure(show="")
            self.show_hide_button.configure(text="Hide")
        else:
            self.entry.configure(show="*")
            self.show_hide_button.configure(text="Show")
    def _ok_event(self, event=None):
        self._password = self.entry.get()
        self.destroy()
    def _cancel_event(self):
        self._password = None
        self.destroy()
    def get_input(self):
        self.wait_window()
        return self._password
    

class BankGUI(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("FinExtract v1.0.8")
        
        window_width = 950
        window_height = 650
        screen_width = self.winfo_screenwidth()
        screen_height = self.winfo_screenheight()
        self.geometry(f'{window_width}x{window_height}+{int(screen_width/2 - window_width/2)}+{int(screen_height/2 - window_height/2)}')
        
        ctk.set_appearance_mode("System")
        self.current_theme_key = "Default (Blue)"
        self.settings_window = None 

        user_home = os.path.expanduser("~")
        self.config_file = os.path.join(user_home, "FinExtract_Settings.json")
        
        self.current_settings = self.load_settings()

        ctk.set_appearance_mode(self.current_settings.get("appearance_mode", "System"))
        self.current_theme_key = self.current_settings.get("theme_name", "Default (Blue)")

        base = os.path.dirname(os.path.abspath(__file__))
        if base not in sys.path: sys.path.insert(0, base)
        self.last_generated_file = None
        self.gui_queue = queue.Queue()
        self.after(100, self.process_gui_queue)

        # Daftar modul untuk Auto-Detect
        self.AUTO_BANKS = [
            ('BNI', 'extract_bni_data', ['BANK NEGARA INDONESIA', 'BNIDIRECT']),
            ('Mandiri', 'process_bank_statement', ['BANK MANDIRI', 'MANDIRI', 'ACCOUNT STATEMENT']),
            ('Livin', 'process_livin_statement', ['LIVIN BY MANDIRI']),
            ('OCBC', 'process_ocbc_final', ['OCBC NISP', 'BANK OCBC', 'OCBC']),
            ('BRI', 'extract_bri_text', ['BANK RAKYAT INDONESIA', 'BRIDIRECT', 'IBBIZ', 'IBIZ', 'BRI'])
        ]

        self.grid_columnconfigure(1, weight=1)
        self.grid_rowconfigure(0, weight=1)

        # --- SIDEBAR ---
        self.sidebar = ctk.CTkFrame(self, width=200, corner_radius=0)
        self.sidebar.grid(row=0, column=0, sticky="nsw")
        self.sidebar.grid_rowconfigure(7, weight=1) 

        self.lbl_sidebar_title = ctk.CTkLabel(self.sidebar, text="Pilih Bank", font=ctk.CTkFont(size=18, weight="bold"))
        self.lbl_sidebar_title.grid(row=0, column=0, padx=20, pady=(20, 10))

        self.btn_bni = ctk.CTkButton(self.sidebar, text="BNI", command=lambda: self.start_task('BNI', 'extract_bni_data'))
        self.btn_bni.grid(row=1, column=0, padx=20, pady=10)
        self.btn_mandiri = ctk.CTkButton(self.sidebar, text="Mandiri", command=lambda: self.start_task('Mandiri', 'process_bank_statement'))
        self.btn_mandiri.grid(row=2, column=0, padx=20, pady=10)
        self.btn_livin = ctk.CTkButton(self.sidebar, text="Livin' by Mandiri", command=lambda: self.start_task('Livin', 'process_livin_statement'))
        self.btn_livin.grid(row=5, column=0, padx=20, pady=10)
        self.btn_ocbc = ctk.CTkButton(self.sidebar, text="OCBC", command=lambda: self.start_task('OCBC', 'process_ocbc_final'))
        self.btn_ocbc.grid(row=3, column=0, padx=20, pady=10)
        self.btn_bri = ctk.CTkButton(self.sidebar, text="BRI", command=lambda: self.start_task('BRI', 'extract_bri_text'))
        self.btn_bri.grid(row=4, column=0, padx=20, pady=10)
        self.btn_auto = ctk.CTkButton(self.sidebar, text="✨ Auto Detect Bank", fg_color="#27ae60", hover_color="#219150", command=lambda: self.start_task('AUTO', ''))
        self.btn_auto.grid(row=6, column=0, padx=20, pady=(20, 10))


        # Settings Button
        self.btn_settings = ctk.CTkButton(self.sidebar, text="⚙ Settings", command=self.open_settings)
        self.btn_settings.grid(row=8, column=0, padx=20, pady=(20, 10), sticky="s")

        self.lbl_credit = ctk.CTkLabel(self.sidebar, text="FinExtract v1.0.8\nBy Michael Aristyo R.", font=ctk.CTkFont(size=10))
        self.lbl_credit.grid(row=9, column=0, padx=10, pady=(0, 10), sticky="s")

        # --- MAIN FRAME ---
        self.main_frame = ctk.CTkFrame(self, corner_radius=0)
        self.main_frame.grid(row=0, column=1, sticky="nsew", padx=(0,0))
        self.main_frame.grid_rowconfigure(1, weight=1) 
        self.main_frame.grid_columnconfigure(0, weight=1)

        # IO Frame
        self.io_frame = ctk.CTkFrame(self.main_frame)
        self.io_frame.grid(row=0, column=0, sticky="ew", padx=20, pady=20)
        self.io_frame.grid_columnconfigure(1, weight=1)

        self.lbl_input = ctk.CTkLabel(self.io_frame, text="Input PDF File(s):")
        self.lbl_input.grid(row=0, column=0, sticky="nw", padx=10, pady=10)
        self.input_textbox = ctk.CTkTextbox(self.io_frame, height=100)
        self.input_textbox.grid(row=0, column=1, sticky="ew", padx=10, pady=10)
        self.input_textbox.configure(state="disabled")
        self.btn_in = ctk.CTkButton(self.io_frame, text="Browse...", width=100, command=self.browse_input)
        self.btn_in.grid(row=0, column=2, sticky="n", padx=10, pady=10)

        self.lbl_output = ctk.CTkLabel(self.io_frame, text="Output Folder:")
        self.lbl_output.grid(row=1, column=0, sticky="w", padx=10, pady=10)
        self.output_entry = ctk.CTkEntry(self.io_frame)
        self.output_entry.grid(row=1, column=1, sticky="ew", padx=10, pady=10)
        self.btn_out = ctk.CTkButton(self.io_frame, text="Browse...", width=100, command=self.browse_output)
        self.btn_out.grid(row=1, column=2, padx=(10, 20), pady=10)

        # Log Frame
        self.log_frame = ctk.CTkFrame(self.main_frame, fg_color="transparent")
        self.log_frame.grid(row=1, column=0, sticky="nsew", padx=20, pady=(0, 0))
        self.log_frame.grid_rowconfigure(1, weight=1)
        self.log_frame.grid_columnconfigure(0, weight=1) 
        
        self.lbl_log_title = ctk.CTkLabel(self.log_frame, text="Log Proses:", font=ctk.CTkFont(weight="bold"))
        self.lbl_log_title.grid(row=0, column=0, sticky="w", padx=10, pady=(10, 5))
        self.btn_clear_log = ctk.CTkButton(self.log_frame, text="Hapus Log", width=100, command=self.clear_log)
        self.btn_clear_log.grid(row=0, column=1, sticky="e", padx=10, pady=(10, 5))

        self.logbox = ctk.CTkTextbox(self.log_frame)
        self.logbox.grid(row=1, column=0, columnspan=2, sticky="nsew", padx=0, pady=(5, 10))
        self.logbox.configure(state="disabled")

        self.logbox.tag_config("SUCCESS", foreground="#00B140") 
        self.logbox.tag_config("ERROR", foreground="#E74C3C")   
        self.logbox.tag_config("PATH", foreground="#3498DB")    
        self.logbox.tag_config("RUN", foreground="#F1C40F")     
        self.logbox.tag_config("SEPARATOR", foreground="#95A5A6")
        self.logbox.tag_config("HEADER", foreground="#00AFFF")

        # Action Frame
        self.action_frame = ctk.CTkFrame(self.main_frame, fg_color="transparent")
        self.action_frame.grid(row=2, column=0, sticky="ew", padx=20, pady=(5, 10))
        self.action_frame.grid_columnconfigure(0, weight=1)
        self.btn_open_folder = ctk.CTkButton(self.action_frame, text="Buka Folder Output", width=150, state="disabled", command=self.open_output_folder)
        self.btn_open_folder.pack(side="right", padx=(5, 0), pady=5)
        self.btn_open_file = ctk.CTkButton(self.action_frame, text="Buka File Output", width=150, state="disabled", command=self.open_output_file)
        self.btn_open_file.pack(side="right", padx=(0, 5), pady=5)

        # Collections
        self.all_buttons = [
            self.btn_bni, self.btn_mandiri, self.btn_ocbc, self.btn_bri, self.btn_livin,
            self.btn_auto, self.btn_settings, 
            self.btn_in, self.btn_out,
            self.btn_clear_log, self.btn_open_folder, self.btn_open_file
        ]
        self.all_labels = [
            self.lbl_sidebar_title, self.lbl_credit,
            self.lbl_input, self.lbl_output, self.lbl_log_title
        ]
        self.all_frames = [self.io_frame]
        self.all_entries = [self.input_textbox, self.output_entry]

        # --- SETUP LOG STARTUP YANG TIDAK BISA DIHAPUS ---
        self.welcome_text = """Selamat Datang di FinExtract!

Fitur Utama:
- Ekstraksi PDF Mutasi Bank ke Excel (BNI, Mandiri, OCBC, BRI, Livin' by Mandiri)
- Auto-detect Password PDF (User input)
- Pengaturan Tema (Settings)

Cara Penggunaan:
1. Pilih File PDF pada panel Input.
2. Tentukan Folder Output.
3. Klik tombol Bank di sidebar kiri untuk memproses.
----------------------------------------------------------------------
"""
        self.logbox.configure(state="normal")
        self.logbox.insert("1.0", self.welcome_text, "HEADER")
        self.logbox.configure(state="disabled")

        self.apply_theme(self.current_theme_key)

    def load_settings(self):
        """Mencoba load file json, jika gagal atau file tidak ada, return default."""
        default_settings = {"appearance_mode": "System", "theme_name": "Default (Blue)"}
        
        try:
            if os.path.exists(self.config_file):
                with open(self.config_file, "r") as f:
                    data = json.load(f)
                
                # Validasi: Pastikan data berhasil di-load dan berupa dictionary
                if data is not None and isinstance(data, dict):
                    return data
            
            # Jika file tidak ada (os.path.exists False), kode akan lanjut ke sini
            
        except Exception as e:
            print(f"Gagal load settings: {e}")
            
        # PENTING: Return default ini akan dieksekusi jika file tidak ada 
        # ATAU terjadi error saat membaca file.
        return default_settings
        
    def save_settings(self):
        """Menyimpan konfigurasi saat ini ke file json."""
        settings = {
            "appearance_mode": ctk.get_appearance_mode(),
            "theme_name": self.current_theme_key
        }
        try:
            with open(self.config_file, "w") as f:
                json.dump(settings, f)
        except Exception as e:
            print(f"Gagal save settings: {e}")
    

    def open_settings(self):
        if self.settings_window is None or not self.settings_window.winfo_exists():
            self.settings_window = SettingsDialog(self)
        else:
            self.settings_window.focus()
            self.settings_window.lift()

    def change_appearance_mode_event(self, new_appearance_mode: str):
        ctk.set_appearance_mode(new_appearance_mode)
        self.after(50, lambda: self.apply_theme(self.current_theme_key))
        self.save_settings()

    def change_theme_event(self, theme_name: str):
        self.apply_theme(theme_name)
        self.save_settings()

    def apply_theme(self, theme_name):
        try:
            self.current_theme_key = theme_name
            theme = THEME_CONFIG.get(theme_name)
            if not theme: return

            self.configure(fg_color=theme["window_bg"])
            self.sidebar.configure(fg_color=theme["sidebar_bg"])
            self.main_frame.configure(fg_color=theme["window_bg"])
            for frame in self.all_frames:
                frame.configure(fg_color=theme["frame_bg"])
            for btn in self.all_buttons:
                btn.configure(fg_color=theme["btn_main"], hover_color=theme["btn_hover"])
            for lbl in self.all_labels:
                lbl.configure(text_color=theme["text"])
            for entry in self.all_entries:
                entry.configure(border_color=theme["border"], fg_color=theme["entry_bg"], text_color=theme["text"])
            
            self.logbox.configure(border_color=theme["border"], fg_color=theme["entry_bg"], text_color=theme["text"])
            
            mode = ctk.get_appearance_mode()
            idx = 0 if mode == "Light" else 1
            self.logbox.tag_config("INFO", foreground=theme["log_info"][idx])
            self.logbox.tag_config("DEFAULT", foreground=theme["text"][idx])

            if self.settings_window and self.settings_window.winfo_exists():
                self.settings_window.update_colors(theme)
            
            # LOG UNTUK PERGANTIAN TEMA DIHAPUS SESUAI PERMINTAAN
            
        except Exception as e:
            print(f"Error applying theme: {e}")

    # --- LOGIKA APLIKASI STANDAR ---
    def browse_input(self):
        paths = filedialog.askopenfilenames(title="Pilih file PDF", filetypes=[("PDF files", "*.pdf")])
        if paths:
            self.input_textbox.configure(state="normal")
            self.input_textbox.delete("1.0", tk.END)
            for path in paths: self.input_textbox.insert(tk.END, path + "\n")
            self.input_textbox.configure(state="disabled")
            self.output_entry.delete(0, tk.END)
            self.output_entry.insert(0, os.path.dirname(paths[0]))

    def browse_output(self):
        path = filedialog.askdirectory()
        if path:
            self.output_entry.delete(0, tk.END)
            self.output_entry.insert(0, path)

    def open_output_folder(self):
        folder_path = self.output_entry.get()
        if not folder_path or not os.path.isdir(folder_path): return
        try:
            if sys.platform == "win32": os.startfile(folder_path)
            elif sys.platform == "darwin": subprocess.Popen(["open", folder_path])
            else: subprocess.Popen(["xdg-open", folder_path])
        except Exception: pass

    def open_output_file(self):
        if self.last_generated_file and os.path.isfile(self.last_generated_file):
            try: os.startfile(self.last_generated_file)
            except Exception: pass

    def log(self, message, level="DEFAULT"):
        try:
            self.logbox.configure(state="normal")
            self.logbox.insert("end", message + "\n", level)
            self.logbox.see("end")
            self.logbox.configure(state="disabled")
        except Exception: print(f"[{level}] {message}")

    def clear_log(self):
        """Hanya menghapus log SETELAH pesan selamat datang."""
        self.logbox.configure(state="normal")
        
        # Hitung baris pesan welcome (termasuk newlines)
        welcome_lines = self.welcome_text.count('\n') + 1 
        # Mulai hapus dari baris setelah welcome message
        delete_start = f"{welcome_lines + 1}.0"
        
        self.logbox.delete(delete_start, tk.END)
        self.logbox.configure(state="disabled")

    def process_gui_queue(self):
        try:
            while True:
                task, args, result_queue = self.gui_queue.get_nowait()
                if task == 'ask_password':
                    self.bell()
                    theme_data = THEME_CONFIG.get(self.current_theme_key)
                    dialog = PasswordDialog(self, title=args['title'], text=args['text'], theme_data=theme_data)
                    self.update_idletasks()
                    dialog.geometry(f"+{self.winfo_x() + (self.winfo_width() - dialog.winfo_reqwidth()) // 2}+{self.winfo_y() + (self.winfo_height() - dialog.winfo_reqheight()) // 2}")                    
                    password = dialog.get_input()
                    result_queue.put(password)
                elif task == 'ask_overwrite':
                    self.bell()
                    response = messagebox.askyesnocancel(**args)
                    result_queue.put(response)
        except queue.Empty: pass
        finally: self.after(100, self.process_gui_queue)

    def start_task(self, module_name, function_name):
        input_paths_str = self.input_textbox.get("1.0", tk.END).strip()
        output_folder = self.output_entry.get().strip()
        if not input_paths_str or not output_folder:
            messagebox.showerror("Error", "Harap pilih file input PDF dan folder output.")
            return
        input_paths = input_paths_str.split('\n')
        self.btn_open_file.configure(state="disabled")
        self.btn_open_folder.configure(state="disabled")
        t = threading.Thread(target=self.process_queue, args=(module_name, function_name, input_paths, output_folder), daemon=True)
        t.start()

    def _request_gui_task(self, task_name, **kwargs):
        result_queue = queue.Queue()
        self.gui_queue.put((task_name, kwargs, result_queue))
        try: return result_queue.get(timeout=300)
        except queue.Empty: return None

    def process_queue(self, module_name, function_name, file_list, output_folder):
        files_processed = 0
        for pdf_path in file_list:
            pdf_path = pdf_path.strip()
            if not pdf_path: continue
            password = None
            try:
                reader = PdfReader(pdf_path) if PdfReader else None
                if reader and getattr(reader, "is_encrypted", False):
                    # Beberapa PDF menyatakan 'encrypted' namun dapat dibuka tanpa password
                    # (owner-only encryption atau password kosong). Coba decrypt dengan
                    # password kosong dahulu; hanya minta input user jika gagal.
                    try:
                        decrypted = False
                        try:
                            res = reader.decrypt("")
                            if res:
                                decrypted = True
                                password = ""  # gunakan empty password
                        except Exception:
                            # Jika decrypt('') tidak didukung atau gagal, lanjutkan
                            pass

                        # Beberapa versi menandai is_encrypted=False setelah decrypt('') berhasil
                        if not decrypted and not getattr(reader, "is_encrypted", False):
                            decrypted = True
                            password = ""

                        if not decrypted:
                            self.log(f"File '{os.path.basename(pdf_path)}' terproteksi.", level="INFO")
                            password = self._request_gui_task('ask_password', text=f"Masukkan password:\n{os.path.basename(pdf_path)}", title="Password")
                            if password is None:
                                self.log("Dibatalkan pengguna.", level="ERROR")
                                continue

                            # Coba decrypt dengan password yg diberikan pengguna
                            try:
                                res = reader.decrypt(password) if reader else False
                                # Jika decrypt gagal, laporkan kesalahan dan lanjut ke file berikutnya
                                if not res and getattr(reader, "is_encrypted", False):
                                    self.log("Password salah.", level="ERROR")
                                    continue
                            except Exception:
                                self.log("Password salah.", level="ERROR")
                                continue
                    except Exception as e:
                        self.log(f"Error saat mencoba decrypt: {e}", level="ERROR")
                        continue
            except Exception as e:
                self.log(f"Error cek enkripsi: {e}", level="ERROR")
                continue

            # Jika kita punya password (termasuk empty string), beberapa modul (pdfplumber/pdfminer)
            # tidak menerima parameter password, maka kita buat salinan PDF yang sudah didekripsi.
            temp_pdf_path = None
            pdf_to_process = pdf_path
            if password is not None and PdfWriter:
                try:
                    reader2 = PdfReader(pdf_path)
                    if reader2.is_encrypted:
                        reader2.decrypt(password)
                    writer = PdfWriter()
                    for p in reader2.pages:
                        writer.add_page(p)
                    tmp = tempfile.NamedTemporaryFile(delete=False, suffix=".pdf")
                    tmp.close()
                    with open(tmp.name, "wb") as f:
                        writer.write(f)
                    temp_pdf_path = tmp.name
                    pdf_to_process = temp_pdf_path
                except Exception as e:
                    self.log(f"Warning: gagal membuat salinan decrypted: {e}", level="ERROR")
            elif password is not None and not PdfWriter:
                self.log("pypdf tidak tersedia; tidak dapat membuat salinan didekripsi. Beberapa file terenkripsi mungkin gagal dibuka oleh pdfplumber.", level="INFO")

            base_name = os.path.basename(pdf_path)
            file_name, _ = os.path.splitext(base_name)
            excel_path = os.path.join(output_folder, f"{file_name}.xlsx")
            if os.path.exists(excel_path):
                response = self._request_gui_task('ask_overwrite', title="Overwrite?", message=f"File ada:\n{excel_path}\nTimpa?")
                if response is None:
                    if temp_pdf_path and os.path.exists(temp_pdf_path): os.remove(temp_pdf_path)
                    return
                elif not response:
                    if temp_pdf_path and os.path.exists(temp_pdf_path): os.remove(temp_pdf_path)
                    continue

            if module_name == 'AUTO':
                # Lapisan 1: Deteksi Keyword untuk Prioritas
                pdf_text = ""
                try:
                    reader_txt = PdfReader(pdf_to_process)
                    if len(reader_txt.pages) > 0:
                        pdf_text = (reader_txt.pages[0].extract_text() or "").upper()
                except: pass

                # Urutkan: yang keyword-nya cocok dicoba duluan
                prioritized = []
                others = []
                for b_mod, b_func, keywords in self.AUTO_BANKS:
                    if any(kw.upper() in pdf_text for kw in keywords): prioritized.append((b_mod, b_func))
                    else: others.append((b_mod, b_func))

                found = False
                for b_mod, b_func in prioritized + others:
                    try:
                        self.log(f"Mencoba format: {b_mod}...", level="INFO")
                        result = self._run_module(b_mod, b_func, pdf_to_process, excel_path, password)
                        if result and isinstance(result, int) and result > 0:
                            self.log(f"Berhasil! Terdeteksi sebagai format {b_mod}.", level="SUCCESS")
                            found = True
                            break
                    except Exception:
                        continue
                if not found:
                    self.log(f"Gagal mendeteksi format untuk: {base_name}", level="ERROR")
            else:
                self._run_module(module_name, function_name, pdf_to_process, excel_path, password)

            if temp_pdf_path and os.path.exists(temp_pdf_path):
                try: os.remove(temp_pdf_path)
                except Exception: pass
            files_processed += 1

        if files_processed > 0:
            self.log("========================================", level="SEPARATOR")
            self.log("SEMUA PROSES SELESAI.\n", level="SUCCESS")
        else:
            self.log("Tidak ada file yang diproses.", level="INFO")

    def _run_module(self, module_name, function_name, pdf_path, excel_path, password=None):
        self.log("-" * 60 + "\n", level="SEPARATOR")
        self.log(f"Memulai: {module_name}", level="HEADER")
        self.log(f"Input: {pdf_path}", level="PATH")
        try:
            mod = importlib.import_module(module_name)
            importlib.reload(mod)
            func = getattr(mod, function_name, None)
            if not func:
                self.log(f"Fungsi '{function_name}' tidak ditemukan.", level="ERROR")
                return
            
            self.log(f"Running...", level="RUN")
            import inspect
            sig = inspect.signature(func)
            call_args = {}
            if 'pdf_path' in sig.parameters: call_args['pdf_path'] = pdf_path
            if 'output_excel' in sig.parameters: call_args['output_excel'] = excel_path
            if 'excel_path' in sig.parameters: call_args['excel_path'] = excel_path
            if 'output_excel_path' in sig.parameters: call_args['output_excel_path'] = excel_path
            if 'password' in sig.parameters and password is not None: call_args['password'] = password

            res = func(**call_args)

            if os.path.exists(excel_path):
                 self.last_generated_file = excel_path
                 self.btn_open_file.configure(state="normal")
                 self.btn_open_folder.configure(state="normal")
            else:
                 self.log("Selesai, tapi file output tidak ditemukan.", level="RUN")
            
            return res
        except Exception as e:
            msg = str(e).lower()
            if "incorrect password" in msg or "not been decrypted" in msg:
                self.log(f"Password salah.", level="ERROR")
            else:
                self.log(f"Error runtime: {e}", level="ERROR")
                self.log(traceback.format_exc(), level="ERROR")

if __name__ == '__main__':
    app = BankGUI()
    app.mainloop()