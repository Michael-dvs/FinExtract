import os
import sys
import json
import shutil
import tempfile
from dataclasses import dataclass, asdict, field
from typing import Dict, Any, Optional, Tuple, Union

from config import SETTINGS_FILE_NAME

SCHEMA_VERSION = 1


@dataclass(eq=False)
class Settings:
    """Model data tunggal untuk seluruh konfigurasi Finextract."""

    schema_version: int = SCHEMA_VERSION

    # Tampilan
    appearance_mode: str = "System"  # "System", "Light", "Dark"
    theme_name: str = "Default (Blue)"  # "Default (Blue)", "Special Edition (Pink)"

    # Format Angka
    thousands_separator: str = ","   # ",", ".", " ", ""
    decimal_separator: str = "."     # ".", ","
    decimal_places: int = 2          # 0, 1, 2, 3, 4
    negative_format: str = "-1234"   # "-1234", "(1234)"
    output_number_type: str = "text" # "text", "numeric"

    # Format Tanggal
    date_format: str = "DD/MM/YYYY"  # "DD/MM/YYYY", "YYYY-MM-DD", "DD-MM-YYYY", "DD MMM YYYY"

    # Output & Penamaan Berkas
    output_folder_mode: str = "last_used"    # "last_used", "ask_every_time"
    last_output_folder: str = ""
    file_naming_pattern: str = "{nama_pdf}"  # "{nama_pdf}", "{nama_pdf}_{bank}", "{nama_pdf}_{tanggal}"
    overwrite_mode: str = "ask"              # "ask", "overwrite", "rename"
    auto_open_output: str = "none"           # "none", "folder", "file"

    # Ekstraksi
    ocr_fallback_enabled: bool = True        # True: OCR aktif saat scraping kosong, False: matikan OCR

    # Privasi & Log
    mask_sensitive_logs: bool = True         # True: samarkan nama/norek/nominal di log
    auto_cleanup_temp_files: bool = True     # True: hapus berkas render sementara

    def get(self, key: str, default: Any = None) -> Any:
        """Kompatibilitas akses dict-like."""
        return getattr(self, key, default)

    def __getitem__(self, key: str) -> Any:
        return getattr(self, key)

    def __eq__(self, other: Any) -> bool:
        if isinstance(other, dict):
            return all(getattr(self, k, None) == v for k, v in other.items())
        if isinstance(other, Settings):
            return asdict(self) == asdict(other)
        return False


    def validate(self) -> Tuple[bool, Optional[str]]:
        """Memvalidasi integritas nilai pengaturan dan aturan silang."""
        # Validasi silang: pemisah ribuan dan desimal tidak boleh sama
        if self.thousands_separator != "" and self.thousands_separator == self.decimal_separator:
            return False, "Pemisah ribuan dan desimal tidak boleh menggunakan karakter yang sama."

        if self.thousands_separator not in (",", ".", " ", ""):
            return False, f"Pemisah ribuan '{self.thousands_separator}' tidak valid."

        if self.decimal_separator not in (".", ","):
            return False, f"Pemisah desimal '{self.decimal_separator}' tidak valid."

        if not isinstance(self.decimal_places, int) or not (0 <= self.decimal_places <= 6):
            return False, "Jumlah digit desimal harus berupa angka antara 0 dan 6."

        if self.negative_format not in ("-1234", "(1234)"):
            return False, f"Format angka negatif '{self.negative_format}' tidak valid."

        if self.output_number_type not in ("text", "numeric"):
            return False, f"Mode keluaran angka '{self.output_number_type}' tidak valid."

        if self.date_format not in ("DD/MM/YYYY", "YYYY-MM-DD", "DD-MM-YYYY", "DD MMM YYYY"):
            return False, f"Format tanggal '{self.date_format}' tidak valid."

        if self.appearance_mode not in ("System", "Light", "Dark"):
            return False, f"Mode tampilan '{self.appearance_mode}' tidak valid."

        if self.theme_name not in ("Default (Blue)", "Special Edition (Pink)"):
            return False, f"Tema '{self.theme_name}' tidak valid."

        if self.overwrite_mode not in ("ask", "overwrite", "rename"):
            return False, f"Perilaku penimpaan file '{self.overwrite_mode}' tidak valid."

        if self.auto_open_output not in ("none", "folder", "file"):
            return False, f"Opsi buka output '{self.auto_open_output}' tidak valid."

        return True, None


class SettingsManager:
    """Pengelola penyimpanan permanen, validasi, penulisan atomik, dan migrasi pengaturan."""

    def __init__(self, config_file: Optional[str] = None):
        if config_file:
            self.config_file = config_file
        else:
            # Utamakan APPDATA di Windows jika tersedia, fallback ke home directory
            appdata = os.environ.get("APPDATA")
            if appdata and sys.platform.startswith("win"):
                fin_dir = os.path.join(appdata, "FinExtract")
                os.makedirs(fin_dir, exist_ok=True)
                self.config_file = os.path.join(fin_dir, SETTINGS_FILE_NAME)
            else:
                self.config_file = os.path.join(os.path.expanduser("~"), SETTINGS_FILE_NAME)

    def load(self) -> Settings:
        """Memuat pengaturan dari berkas JSON.
        
        Tahan rusak:
        - Jika file tidak ada, kembalikan default.
        - Jika file rusak/kosong, buat cadangan (.bak) dan kembalikan default tanpa crash.
        - Jika ada skema lama, lakukan migrasi.
        """
        if not os.path.exists(self.config_file):
            return Settings()

        try:
            with open(self.config_file, "r", encoding="utf-8") as f:
                data = json.load(f)

            if not isinstance(data, dict):
                raise ValueError("Format JSON bukan dictionary.")

            # Migrasi data jika perlu
            migrated = self._migrate(data)

            # Buat objek Settings hanya dengan field yang diakui
            valid_fields = Settings.__dataclass_fields__.keys()
            filtered = {k: v for k, v in migrated.items() if k in valid_fields}
            settings_obj = Settings(**filtered)

            # Validasi nilai
            is_valid, _ = settings_obj.validate()
            if not is_valid:
                # Bila ada nilai korup dalam data, fallback ke default yang valid
                return Settings()

            return settings_obj

        except Exception as e:
            # Cadangkan berkas rusak agar data pengguna tidak hilang permanen
            try:
                bak_file = f"{self.config_file}.bak"
                shutil.copyfile(self.config_file, bak_file)
            except Exception:
                pass
            return Settings()

    def save(self, settings: Union[Settings, Dict[str, Any]]) -> bool:
        """Menyimpan pengaturan ke berkas secara atomik."""
        if isinstance(settings, dict):
            current = self.load()
            valid_fields = Settings.__dataclass_fields__.keys()
            for k, v in settings.items():
                if k in valid_fields:
                    setattr(current, k, v)
            settings = current

        is_valid, err = settings.validate()
        if not is_valid:
            return False

        dirname = os.path.dirname(self.config_file) or "."
        os.makedirs(dirname, exist_ok=True)

        data = asdict(settings)

        try:
            # Penulisan atomik: tulis ke temporary file di direktori yang sama, lalu ganti
            with tempfile.NamedTemporaryFile("w", dir=dirname, delete=False, encoding="utf-8") as tf:
                json.dump(data, tf, indent=2, ensure_ascii=False)
                temp_name = tf.name

            os.replace(temp_name, self.config_file)
            return True
        except Exception:
            if "temp_name" in locals() and os.path.exists(temp_name):
                try:
                    os.remove(temp_name)
                except Exception:
                    pass
            return False

    def export_settings(self, export_path: str, settings: Optional[Settings] = None) -> bool:
        """Mengekspor pengaturan ke berkas JSON tujuan."""
        if settings is None:
            settings = self.load()
        try:
            with open(export_path, "w", encoding="utf-8") as f:
                json.dump(asdict(settings), f, indent=2, ensure_ascii=False)
            return True
        except Exception:
            return False

    def import_settings(self, import_path: str) -> Optional[Settings]:
        """Mengimpor pengaturan dari berkas JSON eksternal."""
        try:
            with open(import_path, "r", encoding="utf-8") as f:
                data = json.load(f)
            migrated = self._migrate(data)
            valid_fields = Settings.__dataclass_fields__.keys()
            filtered = {k: v for k, v in migrated.items() if k in valid_fields}
            new_settings = Settings(**filtered)
            is_valid, _ = new_settings.validate()
            if is_valid:
                self.save(new_settings)
                return new_settings
            return None
        except Exception:
            return None

    def _migrate(self, data: Dict[str, Any]) -> Dict[str, Any]:
        """Menangani peningkatan versi skema (schema migration)."""
        current_version = data.get("schema_version", 0)

        # Versi 0 (format lama di mana hanya ada appearance_mode dan theme_name)
        if current_version < 1:
            data["schema_version"] = 1

        return data

    load_settings = load
    save_settings = save
