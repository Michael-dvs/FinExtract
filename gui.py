import os
import sys
import tkinter as tk
import subprocess
import customtkinter as ctk
from tkinter import filedialog, messagebox
from config import THEME_CONFIG, APP_TITLE
from utils import open_system_path
from settings import Settings, SettingsManager
from formatters import format_number_value, format_date_value, format_output_filename

# --- Mac OS Font Fix ---
if sys.platform == "darwin":
    try:
        _original_ctk_font_init = ctk.CTkFont.__init__
        def _new_ctk_font_init(self, family=None, *args, **kwargs):
            if family is None: family = "SF Pro Display"
            _original_ctk_font_init(self, family, *args, **kwargs)
        ctk.CTkFont.__init__ = _new_ctk_font_init
    except Exception: pass


class SettingsDialog(ctk.CTkToplevel):
    """Jendela Dialog Pengaturan Bergaya Komersial (FinExtract Settings)."""

    def __init__(self, parent, settings_manager=None, save_callback=None):
        super().__init__(parent)
        self.parent = parent
        self.settings_manager = settings_manager or SettingsManager()
        self.save_callback = save_callback

        self.title("Pengaturan - FinExtract")
        self.geometry("760x570")
        self.minsize(740, 540)
        self.resizable(False, False)
        self.attributes("-topmost", True)
        self.focus_force()

        # Posisi jendela di tengah parent
        try:
            parent.update_idletasks()
            pw = parent.winfo_width()
            ph = parent.winfo_height()
            px = parent.winfo_x()
            py = parent.winfo_y()
            x = px + max(0, (pw - 760) // 2)
            y = py + max(0, (ph - 570) // 2)
            self.geometry(f"+{x}+{y}")
        except Exception:
            pass

        self.current_theme = THEME_CONFIG.get(getattr(parent, "current_theme_key", "Default (Blue)"), THEME_CONFIG["Default (Blue)"])
        self.configure(fg_color=self.current_theme["window_bg"])

        self.bind("<Escape>", lambda e: self.destroy())

        # Muat pengaturan aktif
        self.settings: Settings = self.settings_manager.load()

        self._init_mappings()
        self._init_variables()
        self._setup_layout()

        self._select_category("number")
        self._update_live_preview()

    def _init_mappings(self):
        """Memetakan opsi tampilan ke nilai skema internal."""
        self.thousands_map = {
            "Titik (.) [Standar ID/EU]": ".",
            "Koma (,) [Standar US/UK]": ",",
            "Spasi ( )": " ",
            "Tanpa Pemisah": "",
        }
        self.thousands_rev_map = {v: k for k, v in self.thousands_map.items()}

        self.decimal_map = {
            "Koma (,) [Standar ID/EU]": ",",
            "Titik (.) [Standar US/UK]": ".",
        }
        self.decimal_rev_map = {v: k for k, v in self.decimal_map.items()}

        self.neg_map = {
            "Tanda Minus (-1234)": "-1234",
            "Tanda Kurung ((1234))": "(1234)",
        }
        self.neg_rev_map = {v: k for k, v in self.neg_map.items()}

        self.out_type_map = {
            "Teks Terformat (Default - Tampilan Konsisten)": "text",
            "Numerik (Opt-in - Formula/Penjumlahan Excel)": "numeric",
        }
        self.out_type_rev_map = {v: k for k, v in self.out_type_map.items()}

        self.date_map = {
            "DD/MM/YYYY (Contoh: 17/10/2026)": "DD/MM/YYYY",
            "YYYY-MM-DD (Contoh: 2026-10-17)": "YYYY-MM-DD",
            "DD-MM-YYYY (Contoh: 17-10-2026)": "DD-MM-YYYY",
            "DD MMM YYYY (Contoh: 17 Okt 2026)": "DD MMM YYYY",
        }
        self.date_rev_map = {v: k for k, v in self.date_map.items()}

        self.naming_map = {
            "{nama_pdf} (Default)": "{nama_pdf}",
            "{nama_pdf}_{bank}": "{nama_pdf}_{bank}",
            "{nama_pdf}_{tanggal}": "{nama_pdf}_{tanggal}",
            "{bank}_{nama_pdf}": "{bank}_{nama_pdf}",
        }
        self.naming_rev_map = {v: k for k, v in self.naming_map.items()}

        self.overwrite_map = {
            "Tanya Setiap Kali (Konfirmasi)": "ask",
            "Timpa Otomatis (Overwrite)": "overwrite",
            "Beri Nomor Otomatis (file (1).xlsx)": "rename",
        }
        self.overwrite_rev_map = {v: k for k, v in self.overwrite_map.items()}

        self.auto_open_map = {
            "Jangan Buka Apa Pun": "none",
            "Buka File Excel": "file",
            "Buka Folder Output": "folder",
        }
        self.auto_open_rev_map = {v: k for k, v in self.auto_open_map.items()}

    def _init_variables(self):
        """Inisialisasi variabel Tkinter dengan nilai awal dari settings."""
        self.var_thousands = tk.StringVar(value=self.thousands_rev_map.get(self.settings.thousands_separator, "Koma (,) [Standar US/UK]"))
        self.var_decimal = tk.StringVar(value=self.decimal_rev_map.get(self.settings.decimal_separator, "Titik (.) [Standar US/UK]"))
        self.var_dec_places = tk.StringVar(value=str(self.settings.decimal_places))
        self.var_negative = tk.StringVar(value=self.neg_rev_map.get(self.settings.negative_format, "Tanda Minus (-1234)"))
        self.var_out_type = tk.StringVar(value=self.out_type_rev_map.get(self.settings.output_number_type, "Teks Terformat (Default - Tampilan Konsisten)"))

        self.var_date_format = tk.StringVar(value=self.date_rev_map.get(self.settings.date_format, "DD/MM/YYYY (Contoh: 17/10/2026)"))

        self.var_naming = tk.StringVar(value=self.naming_rev_map.get(self.settings.file_naming_pattern, "{nama_pdf} (Default)"))
        self.var_overwrite = tk.StringVar(value=self.overwrite_rev_map.get(self.settings.overwrite_mode, "Tanya Setiap Kali (Konfirmasi)"))
        self.var_auto_open = tk.StringVar(value=self.auto_open_rev_map.get(self.settings.auto_open_output, "Jangan Buka Apa Pun"))

        self.var_ocr = tk.BooleanVar(value=self.settings.ocr_fallback_enabled)
        self.var_mask_logs = tk.BooleanVar(value=self.settings.mask_sensitive_logs)
        self.var_cleanup = tk.BooleanVar(value=self.settings.auto_cleanup_temp_files)

        self.var_mode = tk.StringVar(value=self.settings.appearance_mode)
        self.var_theme = tk.StringVar(value=self.settings.theme_name)

    def _setup_layout(self):
        """Menyusun antarmuka dialog pengaturan (Header, Sidebar, Konten, Preview, Footer)."""
        # --- HEADER ---
        self.header_frame = ctk.CTkFrame(self, fg_color="transparent")
        self.header_frame.pack(fill="x", padx=20, pady=(15, 10))

        self.lbl_title = ctk.CTkLabel(
            self.header_frame,
            text="⚙️  Pengaturan FinExtract",
            font=ctk.CTkFont(size=18, weight="bold"),
            text_color=self.current_theme["text"]
        )
        self.lbl_title.pack(anchor="w")

        self.lbl_subtitle = ctk.CTkLabel(
            self.header_frame,
            text="Konfigurasi format angka, tanggal, alur keluaran berkas, dan privasi.",
            font=ctk.CTkFont(size=11),
            text_color="gray60"
        )
        self.lbl_subtitle.pack(anchor="w", pady=(2, 0))

        # --- BODY (Sidebar + Konten) ---
        self.body_frame = ctk.CTkFrame(self, fg_color="transparent")
        self.body_frame.pack(fill="both", expand=True, padx=20, pady=(0, 5))
        self.body_frame.grid_columnconfigure(1, weight=1)
        self.body_frame.grid_rowconfigure(0, weight=1)

        # Sidebar Kategori
        self.cat_sidebar = ctk.CTkFrame(
            self.body_frame,
            width=180,
            fg_color=self.current_theme["frame_bg"],
            border_width=1,
            border_color=self.current_theme["border"]
        )
        self.cat_sidebar.grid(row=0, column=0, sticky="ns", padx=(0, 10))
        self.cat_sidebar.grid_propagate(False)

        self.cat_buttons = {}
        categories = [
            ("number", "🔢  Format Angka"),
            ("date", "📅  Format Tanggal"),
            ("file", "📁  Berkas & Output"),
            ("extract", "⚙️  Ekstraksi & Privasi"),
            ("appearance", "🎨  Tampilan"),
        ]

        for idx, (cat_id, cat_title) in enumerate(categories):
            btn = ctk.CTkButton(
                self.cat_sidebar,
                text=cat_title,
                anchor="w",
                height=38,
                font=ctk.CTkFont(size=12, weight="bold"),
                fg_color="transparent",
                text_color=self.current_theme["text"],
                hover_color=self.current_theme["btn_hover"],
                command=lambda cid=cat_id: self._select_category(cid)
            )
            btn.pack(fill="x", padx=8, pady=4)
            self.cat_buttons[cat_id] = btn

        # Container Konten Kanan
        self.content_container = ctk.CTkFrame(
            self.body_frame,
            fg_color=self.current_theme["frame_bg"],
            border_width=1,
            border_color=self.current_theme["border"]
        )
        self.content_container.grid(row=0, column=1, sticky="nsew")

        # Live Preview Frame (di bagian bawah konten)
        self.preview_frame = ctk.CTkFrame(
            self.content_container,
            fg_color=self.current_theme["entry_bg"],
            border_width=1,
            border_color=self.current_theme["border"]
        )
        self.preview_frame.pack(fill="x", padx=15, pady=(10, 8), side="bottom")

        self.lbl_preview_title = ctk.CTkLabel(
            self.preview_frame,
            text="🔍  Pratinjau Langsung (Live Preview):",
            font=ctk.CTkFont(size=12, weight="bold"),
            text_color=self.current_theme["text"]
        )
        self.lbl_preview_title.pack(anchor="w", padx=12, pady=(6, 2))

        self.lbl_preview_text = ctk.CTkLabel(
            self.preview_frame,
            text="",
            font=ctk.CTkFont(family="Consolas" if sys.platform != "darwin" else "Menlo", size=11),
            text_color=self.current_theme["text"],
            justify="left"
        )
        self.lbl_preview_text.pack(anchor="w", padx=12, pady=(2, 6))

        # Label Validasi Error
        self.lbl_validation_err = ctk.CTkLabel(
            self.content_container,
            text="",
            font=ctk.CTkFont(size=11, weight="bold"),
            text_color="#E74C3C"
        )
        self.lbl_validation_err.pack(anchor="w", padx=15, pady=(0, 2), side="bottom")

        # Frame Panel per Kategori
        self.cat_frames = {}
        self._build_number_tab()
        self._build_date_tab()
        self._build_file_tab()
        self._build_extract_tab()
        self._build_appearance_tab()

        # --- FOOTER ---
        self.footer_frame = ctk.CTkFrame(self, fg_color="transparent")
        self.footer_frame.pack(fill="x", padx=20, pady=(6, 15), side="bottom")

        # Tombol Kiri (Reset, Ekspor, Impor)
        self.btn_reset = ctk.CTkButton(
            self.footer_frame,
            text="🔄 Reset Default",
            width=115,
            fg_color="transparent",
            border_width=1,
            border_color=self.current_theme["border"],
            text_color=self.current_theme["text"],
            hover_color=self.current_theme["btn_hover"],
            command=self.on_reset
        )
        self.btn_reset.pack(side="left", padx=(0, 6))

        self.btn_export = ctk.CTkButton(
            self.footer_frame,
            text="📤 Ekspor...",
            width=90,
            fg_color="transparent",
            border_width=1,
            border_color=self.current_theme["border"],
            text_color=self.current_theme["text"],
            hover_color=self.current_theme["btn_hover"],
            command=self.on_export
        )
        self.btn_export.pack(side="left", padx=4)

        self.btn_import = ctk.CTkButton(
            self.footer_frame,
            text="📥 Impor...",
            width=90,
            fg_color="transparent",
            border_width=1,
            border_color=self.current_theme["border"],
            text_color=self.current_theme["text"],
            hover_color=self.current_theme["btn_hover"],
            command=self.on_import
        )
        self.btn_import.pack(side="left", padx=4)

        # Tombol Kanan (Batal, Simpan)
        self.btn_save = ctk.CTkButton(
            self.footer_frame,
            text="💾 Simpan",
            width=100,
            font=ctk.CTkFont(weight="bold"),
            fg_color=self.current_theme["btn_main"],
            hover_color=self.current_theme["btn_hover"],
            command=self.on_save
        )
        self.btn_save.pack(side="right", padx=(6, 0))

        self.btn_cancel = ctk.CTkButton(
            self.footer_frame,
            text="Batal",
            width=85,
            fg_color="gray50",
            hover_color="gray40",
            command=self.destroy
        )
        self.btn_cancel.pack(side="right", padx=4)

    def _select_category(self, cat_id: str):
        """Beralih ke panel kategori yang dipilih."""
        for cid, btn in self.cat_buttons.items():
            if cid == cat_id:
                btn.configure(fg_color=self.current_theme["btn_main"], text_color="white")
            else:
                btn.configure(fg_color="transparent", text_color=self.current_theme["text"])

        for cid, frame in self.cat_frames.items():
            if cid == cat_id:
                frame.pack(fill="both", expand=True, padx=15, pady=10)
            else:
                frame.pack_forget()

    def _build_number_tab(self):
        """Panel Pengaturan Format Angka."""
        frame = ctk.CTkFrame(self.content_container, fg_color="transparent")
        self.cat_frames["number"] = frame

        # Pemisah Ribuan
        lbl_th = ctk.CTkLabel(frame, text="Pemisah Ribuan:", font=ctk.CTkFont(weight="bold"), text_color=self.current_theme["text"])
        lbl_th.pack(anchor="w", pady=(5, 2))
        menu_th = ctk.CTkOptionMenu(
            frame,
            values=list(self.thousands_map.keys()),
            variable=self.var_thousands,
            command=self._update_live_preview,
            fg_color=self.current_theme["btn_main"],
            button_color=self.current_theme["btn_hover"],
            width=280
        )
        menu_th.pack(anchor="w", pady=(0, 8))

        # Pemisah Desimal
        lbl_dec = ctk.CTkLabel(frame, text="Pemisah Desimal:", font=ctk.CTkFont(weight="bold"), text_color=self.current_theme["text"])
        lbl_dec.pack(anchor="w", pady=(2, 2))
        menu_dec = ctk.CTkOptionMenu(
            frame,
            values=list(self.decimal_map.keys()),
            variable=self.var_decimal,
            command=self._update_live_preview,
            fg_color=self.current_theme["btn_main"],
            button_color=self.current_theme["btn_hover"],
            width=280
        )
        menu_dec.pack(anchor="w", pady=(0, 8))

        # Digit Desimal
        lbl_dp = ctk.CTkLabel(frame, text="Jumlah Digit Desimal:", font=ctk.CTkFont(weight="bold"), text_color=self.current_theme["text"])
        lbl_dp.pack(anchor="w", pady=(2, 2))
        menu_dp = ctk.CTkOptionMenu(
            frame,
            values=["0", "1", "2", "3", "4"],
            variable=self.var_dec_places,
            command=self._update_live_preview,
            fg_color=self.current_theme["btn_main"],
            button_color=self.current_theme["btn_hover"],
            width=120
        )
        menu_dp.pack(anchor="w", pady=(0, 8))

        # Format Negatif
        lbl_neg = ctk.CTkLabel(frame, text="Format Angka Negatif:", font=ctk.CTkFont(weight="bold"), text_color=self.current_theme["text"])
        lbl_neg.pack(anchor="w", pady=(2, 2))
        menu_neg = ctk.CTkOptionMenu(
            frame,
            values=list(self.neg_map.keys()),
            variable=self.var_negative,
            command=self._update_live_preview,
            fg_color=self.current_theme["btn_main"],
            button_color=self.current_theme["btn_hover"],
            width=280
        )
        menu_neg.pack(anchor="w", pady=(0, 8))

        # Mode Keluaran Excel
        lbl_out = ctk.CTkLabel(frame, text="Mode Keluaran Angka di Excel:", font=ctk.CTkFont(weight="bold"), text_color=self.current_theme["text"])
        lbl_out.pack(anchor="w", pady=(2, 2))
        menu_out = ctk.CTkOptionMenu(
            frame,
            values=list(self.out_type_map.keys()),
            variable=self.var_out_type,
            command=self._update_live_preview,
            fg_color=self.current_theme["btn_main"],
            button_color=self.current_theme["btn_hover"],
            width=360
        )
        menu_out.pack(anchor="w", pady=(0, 4))

        lbl_out_hint = ctk.CTkLabel(
            frame,
            text="*Teks menjamin tampilan pemisah konsisten di semua regional komputer.\n*Numerik memungkinkan penggunaan rumus SUM/perhitungan langsung di Excel.",
            font=ctk.CTkFont(size=10),
            text_color="gray60",
            justify="left"
        )
        lbl_out_hint.pack(anchor="w", pady=(0, 5))

    def _build_date_tab(self):
        """Panel Pengaturan Format Tanggal."""
        frame = ctk.CTkFrame(self.content_container, fg_color="transparent")
        self.cat_frames["date"] = frame

        lbl_dt = ctk.CTkLabel(frame, text="Pola Format Tanggal:", font=ctk.CTkFont(weight="bold"), text_color=self.current_theme["text"])
        lbl_dt.pack(anchor="w", pady=(10, 4))

        menu_dt = ctk.CTkOptionMenu(
            frame,
            values=list(self.date_map.keys()),
            variable=self.var_date_format,
            command=self._update_live_preview,
            fg_color=self.current_theme["btn_main"],
            button_color=self.current_theme["btn_hover"],
            width=320
        )
        menu_dt.pack(anchor="w", pady=(0, 10))

        lbl_hint = ctk.CTkLabel(
            frame,
            text="Format tanggal diterapkan ke kolom tanggal (Posting Date / Tanggal Transaksi)\npada mutasi rekening Excel yang dihasilkan.",
            font=ctk.CTkFont(size=11),
            text_color="gray60",
            justify="left"
        )
        lbl_hint.pack(anchor="w", pady=(5, 10))

    def _build_file_tab(self):
        """Panel Pengaturan Berkas & Output."""
        frame = ctk.CTkFrame(self.content_container, fg_color="transparent")
        self.cat_frames["file"] = frame

        # Pola Nama Berkas
        lbl_nm = ctk.CTkLabel(frame, text="Pola Penamaan Berkas Excel:", font=ctk.CTkFont(weight="bold"), text_color=self.current_theme["text"])
        lbl_nm.pack(anchor="w", pady=(10, 2))
        menu_nm = ctk.CTkOptionMenu(
            frame,
            values=list(self.naming_map.keys()),
            variable=self.var_naming,
            command=self._update_live_preview,
            fg_color=self.current_theme["btn_main"],
            button_color=self.current_theme["btn_hover"],
            width=280
        )
        menu_nm.pack(anchor="w", pady=(0, 10))

        # Perilaku File Kembar
        lbl_ow = ctk.CTkLabel(frame, text="Perilaku Jika Berkas Tujuan Sudah Ada:", font=ctk.CTkFont(weight="bold"), text_color=self.current_theme["text"])
        lbl_ow.pack(anchor="w", pady=(5, 2))
        menu_ow = ctk.CTkOptionMenu(
            frame,
            values=list(self.overwrite_map.keys()),
            variable=self.var_overwrite,
            fg_color=self.current_theme["btn_main"],
            button_color=self.current_theme["btn_hover"],
            width=280
        )
        menu_ow.pack(anchor="w", pady=(0, 10))

        # Buka Otomatis
        lbl_ao = ctk.CTkLabel(frame, text="Buka Otomatis Setelah Ekstraksi Selesai:", font=ctk.CTkFont(weight="bold"), text_color=self.current_theme["text"])
        lbl_ao.pack(anchor="w", pady=(5, 2))
        menu_ao = ctk.CTkOptionMenu(
            frame,
            values=list(self.auto_open_map.keys()),
            variable=self.var_auto_open,
            fg_color=self.current_theme["btn_main"],
            button_color=self.current_theme["btn_hover"],
            width=280
        )
        menu_ao.pack(anchor="w", pady=(0, 10))

    def _build_extract_tab(self):
        """Panel Pengaturan Ekstraksi & Privasi."""
        frame = ctk.CTkFrame(self.content_container, fg_color="transparent")
        self.cat_frames["extract"] = frame

        # Switch OCR
        sw_ocr = ctk.CTkSwitch(
            frame,
            text="Aktifkan Fallback OCR saat PDF scan/0 data",
            variable=self.var_ocr,
            progress_color=self.current_theme["btn_main"]
        )
        sw_ocr.pack(anchor="w", pady=(12, 2))
        lbl_ocr_hint = ctk.CTkLabel(
            frame,
            text="Bila aktif, aplikasi otomatis beralih ke OCR jika PDF tidak memuat teks teks digital.",
            font=ctk.CTkFont(size=10),
            text_color="gray60"
        )
        lbl_ocr_hint.pack(anchor="w", padx=25, pady=(0, 12))

        # Switch Masking Log
        sw_mask = ctk.CTkSwitch(
            frame,
            text="Samarkan Data Sensitif di Log UI",
            variable=self.var_mask_logs,
            progress_color=self.current_theme["btn_main"]
        )
        sw_mask.pack(anchor="w", pady=(6, 2))
        lbl_mask_hint = ctk.CTkLabel(
            frame,
            text="Menyamarkan nomor rekening dan nominal saldo di kotak log demi kerahasiaan.",
            font=ctk.CTkFont(size=10),
            text_color="gray60"
        )
        lbl_mask_hint.pack(anchor="w", padx=25, pady=(0, 12))

        # Switch Pembersihan Temp
        sw_clean = ctk.CTkSwitch(
            frame,
            text="Hapus Otomatis Berkas Render Gambar OCR Sementara",
            variable=self.var_cleanup,
            progress_color=self.current_theme["btn_main"]
        )
        sw_clean.pack(anchor="w", pady=(6, 2))
        lbl_clean_hint = ctk.CTkLabel(
            frame,
            text="Menghapus gambar render halaman setelah ekstraksi selesai untuk menghemat memori.",
            font=ctk.CTkFont(size=10),
            text_color="gray60"
        )
        lbl_clean_hint.pack(anchor="w", padx=25, pady=(0, 10))

    def _build_appearance_tab(self):
        """Panel Pengaturan Tampilan & Tema."""
        frame = ctk.CTkFrame(self.content_container, fg_color="transparent")
        self.cat_frames["appearance"] = frame

        # Mode Tampilan
        lbl_mode = ctk.CTkLabel(frame, text="Mode Tampilan:", font=ctk.CTkFont(weight="bold"), text_color=self.current_theme["text"])
        lbl_mode.pack(anchor="w", pady=(10, 4))
        seg_mode = ctk.CTkSegmentedButton(
            frame,
            values=["System", "Light", "Dark"],
            variable=self.var_mode,
            command=self._on_mode_change,
            selected_color=self.current_theme["btn_main"],
            selected_hover_color=self.current_theme["btn_hover"]
        )
        seg_mode.pack(anchor="w", pady=(0, 12))

        # Tema Warna
        lbl_theme = ctk.CTkLabel(frame, text="Tema Warna:", font=ctk.CTkFont(weight="bold"), text_color=self.current_theme["text"])
        lbl_theme.pack(anchor="w", pady=(6, 4))
        menu_theme = ctk.CTkOptionMenu(
            frame,
            values=["Default (Blue)", "Special Edition (Pink)"],
            variable=self.var_theme,
            command=self._on_theme_change,
            fg_color=self.current_theme["btn_main"],
            button_color=self.current_theme["btn_hover"],
            width=240
        )
        menu_theme.pack(anchor="w", pady=(0, 10))

    def _on_mode_change(self, value=None):
        mode = self.var_mode.get()
        if self.parent and hasattr(self.parent, "change_appearance_mode_event"):
            self.parent.change_appearance_mode_event(mode)

    def _on_theme_change(self, value=None):
        theme_key = self.var_theme.get()
        if self.parent and hasattr(self.parent, "change_theme_event"):
            self.parent.change_theme_event(theme_key)

    def _update_live_preview(self, *args):
        """Memperbarui teks pratinjau langsung berdasarkan pengaturan yang sedang dipilih."""
        th_sep = self.thousands_map.get(self.var_thousands.get(), ",")
        dec_sep = self.decimal_map.get(self.var_decimal.get(), ".")
        try:
            dp = int(self.var_dec_places.get())
        except ValueError:
            dp = 2
        neg_fmt = self.neg_map.get(self.var_negative.get(), "-1234")
        out_type = self.out_type_map.get(self.var_out_type.get(), "text")
        dt_fmt = self.date_map.get(self.var_date_format.get(), "DD/MM/YYYY")

        # Validasi silang pemisah ribuan dan desimal
        if th_sep != "" and th_sep == dec_sep:
            self.lbl_validation_err.configure(
                text="⚠️ Peringatan: Pemisah ribuan dan desimal tidak boleh menggunakan karakter yang sama!"
            )
            self.btn_save.configure(state="disabled")
        else:
            self.lbl_validation_err.configure(text="")
            self.btn_save.configure(state="normal")

        # Buat objek Settings sementara untuk memformat contoh
        sample_s = Settings(
            thousands_separator=th_sep,
            decimal_separator=dec_sep,
            decimal_places=dp,
            negative_format=neg_fmt,
            output_number_type=out_type,
            date_format=dt_fmt,
        )

        pos_str = format_number_value("1234567.89", sample_s)
        neg_str = format_number_value("-1234.50", sample_s)
        dt_str = format_date_value("17/10/2026", dt_fmt)
        naming_pattern = self.naming_map.get(self.var_naming.get(), "{nama_pdf}")
        file_str = format_output_filename(naming_pattern, "Mandiri_2026", "Mandiri")

        preview_msg = (
            f"Positif: {pos_str}    |    Negatif: {neg_str}    |    Tanggal: {dt_str}\n"
            f"Nama Berkas: {file_str}.xlsx    |    Mode Excel: {out_type.upper()}"
        )
        self.lbl_preview_text.configure(text=preview_msg)

    def _build_settings_from_form(self) -> Settings:
        """Mengumpulkan nilai dari antarmuka formulir menjadi objek Settings."""
        th_sep = self.thousands_map.get(self.var_thousands.get(), ",")
        dec_sep = self.decimal_map.get(self.var_decimal.get(), ".")
        try:
            dp = int(self.var_dec_places.get())
        except ValueError:
            dp = 2
        neg_fmt = self.neg_map.get(self.var_negative.get(), "-1234")
        out_type = self.out_type_map.get(self.var_out_type.get(), "text")
        dt_fmt = self.date_map.get(self.var_date_format.get(), "DD/MM/YYYY")
        naming = self.naming_map.get(self.var_naming.get(), "{nama_pdf}")
        overwrite = self.overwrite_map.get(self.var_overwrite.get(), "ask")
        auto_open = self.auto_open_map.get(self.var_auto_open.get(), "none")

        return Settings(
            appearance_mode=self.var_mode.get(),
            theme_name=self.var_theme.get(),
            thousands_separator=th_sep,
            decimal_separator=dec_sep,
            decimal_places=dp,
            negative_format=neg_fmt,
            output_number_type=out_type,
            date_format=dt_fmt,
            file_naming_pattern=naming,
            overwrite_mode=overwrite,
            auto_open_output=auto_open,
            ocr_fallback_enabled=self.var_ocr.get(),
            mask_sensitive_logs=self.var_mask_logs.get(),
            auto_cleanup_temp_files=self.var_cleanup.get(),
        )

    def on_save(self):
        """Menyimpan seluruh pengaturan ke file secara atomik dan menutup dialog."""
        new_settings = self._build_settings_from_form()
        is_valid, err = new_settings.validate()
        if not is_valid:
            messagebox.showerror("Pengaturan Tidak Valid", err, parent=self)
            return

        success = self.settings_manager.save(new_settings)
        if not success:
            messagebox.showerror("Error", "Gagal menyimpan pengaturan ke berkas konfigurasi.", parent=self)
            return

        # Terapkan perubahan tema/mode
        if self.parent:
            if hasattr(self.parent, "change_appearance_mode_event"):
                self.parent.change_appearance_mode_event(new_settings.appearance_mode)
            if hasattr(self.parent, "change_theme_event"):
                self.parent.change_theme_event(new_settings.theme_name)
            if self.save_callback:
                try:
                    self.save_callback()
                except Exception:
                    pass

        self.destroy()

    def on_reset(self):
        """Mengembalikan seluruh isian form ke nilai default pabrik."""
        if not messagebox.askyesno("Konfirmasi Reset", "Kembalikan seluruh pengaturan ke nilai default?", parent=self):
            return
        default = Settings()
        self._load_settings_into_form(default)
        self._update_live_preview()

    def on_export(self):
        """Mengekspor pengaturan saat ini ke file JSON."""
        export_path = filedialog.asksaveasfilename(
            parent=self,
            title="Ekspor Pengaturan ke JSON",
            defaultextension=".json",
            filetypes=[("JSON Files", "*.json")],
            initialfile="finextract_settings.json"
        )
        if not export_path:
            return

        current_draft = self._build_settings_from_form()
        if self.settings_manager.export_settings(export_path, current_draft):
            messagebox.showinfo("Sukses", f"Pengaturan berhasil diekspor ke:\n{export_path}", parent=self)
        else:
            messagebox.showerror("Error", "Gagal mengekspor berkas pengaturan.", parent=self)

    def on_import(self):
        """Mengimpor pengaturan dari file JSON eksternal."""
        import_path = filedialog.askopenfilename(
            parent=self,
            title="Impor Pengaturan dari JSON",
            filetypes=[("JSON Files", "*.json")]
        )
        if not import_path:
            return

        imported = self.settings_manager.import_settings(import_path)
        if imported:
            self.settings = imported
            self._load_settings_into_form(imported)
            self._update_live_preview()
            if self.parent:
                if hasattr(self.parent, "change_appearance_mode_event"):
                    self.parent.change_appearance_mode_event(imported.appearance_mode)
                if hasattr(self.parent, "change_theme_event"):
                    self.parent.change_theme_event(imported.theme_name)
            messagebox.showinfo("Sukses", "Pengaturan berhasil diimpor dan disimpan.", parent=self)
        else:
            messagebox.showerror("Error", "Berkas JSON tidak valid atau gagal diimpor.", parent=self)

    def _load_settings_into_form(self, s: Settings):
        """Memuat objek Settings ke variabel form antarmuka."""
        self.var_thousands.set(self.thousands_rev_map.get(s.thousands_separator, ""))
        self.var_decimal.set(self.decimal_rev_map.get(s.decimal_separator, ""))
        self.var_dec_places.set(str(s.decimal_places))
        self.var_negative.set(self.neg_rev_map.get(s.negative_format, ""))
        self.var_out_type.set(self.out_type_rev_map.get(s.output_number_type, ""))
        self.var_date_format.set(self.date_rev_map.get(s.date_format, ""))
        self.var_naming.set(self.naming_rev_map.get(s.file_naming_pattern, ""))
        self.var_overwrite.set(self.overwrite_rev_map.get(s.overwrite_mode, ""))
        self.var_auto_open.set(self.auto_open_rev_map.get(s.auto_open_output, ""))
        self.var_ocr.set(s.ocr_fallback_enabled)
        self.var_mask_logs.set(s.mask_sensitive_logs)
        self.var_cleanup.set(s.auto_cleanup_temp_files)
        self.var_mode.set(s.appearance_mode)
        self.var_theme.set(s.theme_name)

    def update_colors(self, theme_config):
        """Memperbarui palet warna seluruh widget jika tema berubah saat dialog terbuka."""
        try:
            self.current_theme = theme_config
            self.configure(fg_color=theme_config["window_bg"])
            self.lbl_title.configure(text_color=theme_config["text"])
            self.cat_sidebar.configure(fg_color=theme_config["frame_bg"], border_color=theme_config["border"])
            self.content_container.configure(fg_color=theme_config["frame_bg"], border_color=theme_config["border"])
            self.preview_frame.configure(fg_color=theme_config["entry_bg"], border_color=theme_config["border"])
            self.lbl_preview_title.configure(text_color=theme_config["text"])
            self.lbl_preview_text.configure(text_color=theme_config["text"])
            self.btn_save.configure(fg_color=theme_config["btn_main"], hover_color=theme_config["btn_hover"])
        except Exception:
            pass

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
        self._password = None

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

class FinextractUI(ctk.CTk):
    def __init__(self, settings_callback=None, settings_manager=None):
        super().__init__()
        self.title(APP_TITLE)
        
        window_width = 950
        window_height = 650
        screen_width = self.winfo_screenwidth()
        screen_height = self.winfo_screenheight()
        self.geometry(f'{window_width}x{window_height}+{int(screen_width/2 - window_width/2)}+{int(screen_height/2 - window_height/2)}')
        
        self.settings_manager = settings_manager or SettingsManager()
        self.settings_window = None 
        self.current_theme_key = "Default (Blue)"
        self.save_settings_callback = settings_callback
        self.process_callback = None # Function to call when bank button is clicked
        
        self.last_generated_file = None

        # Pintasan Keyboard Pengaturan (Ctrl+, dan Cmd+,)
        self.bind("<Control-comma>", lambda e: self.open_settings())
        self.bind("<Command-comma>", lambda e: self.open_settings())

        self.grid_columnconfigure(1, weight=1)
        self.grid_rowconfigure(0, weight=1)

        self._setup_sidebar()
        self._setup_main_frame()
        
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

    def _setup_sidebar(self):
        self.sidebar = ctk.CTkFrame(self, width=200, corner_radius=0)
        self.sidebar.grid(row=0, column=0, sticky="nsw")
        self.sidebar.grid_rowconfigure(7, weight=1) 

        self.lbl_sidebar_title = ctk.CTkLabel(self.sidebar, text="Pilih Bank", font=ctk.CTkFont(size=18, weight="bold"))
        self.lbl_sidebar_title.grid(row=0, column=0, padx=20, pady=(20, 10))

        # Buttons
        self.btn_bni = ctk.CTkButton(self.sidebar, text="BNI", command=lambda: self.trigger_process('parser.BNI', 'extract_bni_data'))
        self.btn_bni.grid(row=1, column=0, padx=20, pady=10)
        self.btn_mandiri = ctk.CTkButton(self.sidebar, text="Mandiri", command=lambda: self.trigger_process('parser.Mandiri', 'process_bank_statement'))
        self.btn_mandiri.grid(row=2, column=0, padx=20, pady=10)
        self.btn_livin = ctk.CTkButton(self.sidebar, text="Livin' by Mandiri", command=lambda: self.trigger_process('parser.Livin', 'process_livin_statement'))
        self.btn_livin.grid(row=5, column=0, padx=20, pady=10)
        self.btn_ocbc = ctk.CTkButton(self.sidebar, text="OCBC", command=lambda: self.trigger_process('parser.OCBC', 'process_ocbc_final'))
        self.btn_ocbc.grid(row=3, column=0, padx=20, pady=10)
        self.btn_bri = ctk.CTkButton(self.sidebar, text="BRI", command=lambda: self.trigger_process('parser.BRI', 'extract_bri_text'))
        self.btn_bri.grid(row=4, column=0, padx=20, pady=10)
        self.btn_auto = ctk.CTkButton(self.sidebar, text="✨ Auto Detect Bank", fg_color="#27ae60", hover_color="#219150", command=lambda: self.trigger_process('AUTO', ''))
        self.btn_auto.grid(row=6, column=0, padx=20, pady=(20, 5))

        self.lbl_fallback_info = ctk.CTkLabel(
            self.sidebar,
            text="*PDF scan otomatis via OCR\n(Mandiri didukung; bank lain [DATA BELUM TERSEDIA])",
            font=ctk.CTkFont(size=9),
            text_color="gray60",
            justify="center"
        )
        self.lbl_fallback_info.grid(row=7, column=0, padx=10, pady=(0, 10))

        self.sidebar.grid_rowconfigure(8, weight=1) 

        self.btn_settings = ctk.CTkButton(self.sidebar, text="⚙ Pengaturan (Ctrl+,)", command=self.open_settings)
        self.btn_settings.grid(row=9, column=0, padx=20, pady=(20, 10), sticky="s")

        self.lbl_credit = ctk.CTkLabel(self.sidebar, text=f"{APP_TITLE}\nBy Michael Aristyo R.", font=ctk.CTkFont(size=10))
        self.lbl_credit.grid(row=10, column=0, padx=10, pady=(0, 10), sticky="s")

    def _setup_main_frame(self):
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
        self.logbox.tag_config("WARNING", foreground="#E67E22")
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

        # Collections for theming
        self.all_buttons = [
            self.btn_bni, self.btn_mandiri, self.btn_ocbc, self.btn_bri, self.btn_livin,
            self.btn_auto, self.btn_settings, 
            self.btn_in, self.btn_out,
            self.btn_clear_log, self.btn_open_folder, self.btn_open_file
        ]
        self.all_labels = [
            self.lbl_sidebar_title, self.lbl_credit,
            self.lbl_input, self.lbl_output, self.lbl_log_title, self.lbl_fallback_info
        ]
        self.all_frames = [self.io_frame]
        self.all_entries = [self.input_textbox, self.output_entry]

    # --- UI Interaction Methods ---
    def set_process_callback(self, callback):
        self.process_callback = callback

    def trigger_process(self, module, func):
        if self.process_callback:
            self.process_callback(module, func)
        else:
            messagebox.showerror("Error", "Logic controller not connected.")

    def get_input_files(self):
        text = self.input_textbox.get("1.0", tk.END).strip()
        return text.split('\n') if text else []

    def get_output_folder(self):
        return self.output_entry.get().strip()

    def enable_open_buttons(self, file_path=None):
        if file_path:
            self.last_generated_file = file_path
            self.btn_open_file.configure(state="normal")
        self.btn_open_folder.configure(state="normal")

    def disable_open_buttons(self):
        self.btn_open_file.configure(state="disabled")
        self.btn_open_folder.configure(state="disabled")

    def log_message(self, message, level="DEFAULT"):
        try:
            self.logbox.configure(state="normal")
            self.logbox.insert("end", message + "\n", level)
            self.logbox.see("end")
            self.logbox.configure(state="disabled")
        except Exception: pass

    def ask_password(self, title, text):
        self.bell()
        theme_data = THEME_CONFIG.get(self.current_theme_key)
        dialog = PasswordDialog(self, title=title, text=text, theme_data=theme_data)
        self.update_idletasks()
        dialog.geometry(f"+{self.winfo_x() + (self.winfo_width() - dialog.winfo_reqwidth()) // 2}+{self.winfo_y() + (self.winfo_height() - dialog.winfo_reqheight()) // 2}")                    
        return dialog.get_input()

    def ask_overwrite(self, title, message):
        self.bell()
        return messagebox.askyesnocancel(title=title, message=message)

    # --- Standard UI Methods ---
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
        if folder_path and os.path.isdir(folder_path):
            open_system_path(folder_path)

    def open_output_file(self):
        if self.last_generated_file and os.path.isfile(self.last_generated_file):
            open_system_path(self.last_generated_file)

    def clear_log(self):
        self.logbox.configure(state="normal")
        welcome_lines = self.welcome_text.count('\n') + 1 
        delete_start = f"{welcome_lines + 1}.0"
        self.logbox.delete(delete_start, tk.END)
        self.logbox.configure(state="disabled")

    def open_settings(self):
        if self.settings_window is None or not self.settings_window.winfo_exists():
            self.settings_window = SettingsDialog(
                self,
                settings_manager=self.settings_manager,
                save_callback=self.save_settings_callback,
            )
        else:
            self.settings_window.focus()
            self.settings_window.lift()

    def change_appearance_mode_event(self, new_appearance_mode: str):
        ctk.set_appearance_mode(new_appearance_mode)
        self.after(50, lambda: self.apply_theme(self.current_theme_key))
        if self.save_settings_callback: self.save_settings_callback()

    def change_theme_event(self, theme_name: str):
        self.apply_theme(theme_name)
        if self.save_settings_callback: self.save_settings_callback()

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
        except Exception as e:
            print(f"Error applying theme: {e}")
