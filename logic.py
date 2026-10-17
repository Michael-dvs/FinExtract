import os
import json
import threading
import importlib
import tempfile
import inspect
import queue
from typing import Any, Dict, List, Optional, Tuple
from config import AUTO_BANKS, SETTINGS_FILE_NAME


class CoreLogic:
    """Mengelola alur pemrosesan ekstraksi PDF, penanganan enkripsi,

    dan koordinasi komunikasi dengan antarmuka GUI.
    """

    def __init__(self, status_queue: queue.Queue, request_queue: queue.Queue):
        self.status_queue = status_queue
        self.request_queue = request_queue
        self.config_file = os.path.join(os.path.expanduser("~"), SETTINGS_FILE_NAME)
        self.AUTO_BANKS = AUTO_BANKS

    def load_settings(self) -> Dict[str, Any]:
        """Membaca pengaturan tampilan dari file json di direktori pengguna."""
        default = {"appearance_mode": "System", "theme_name": "Default (Blue)"}
        try:
            if os.path.exists(self.config_file):
                with open(self.config_file, "r") as f:
                    return json.load(f)
        except Exception:
            pass
        return default

    def save_settings(self, settings: Dict[str, Any]) -> None:
        """Menyimpan pengaturan tampilan pengguna."""
        try:
            with open(self.config_file, "w") as f:
                json.dump(settings, f)
        except Exception:
            pass

    def start_processing_thread(
        self,
        module_name: str,
        function_name: str,
        file_list: List[str],
        output_folder: str,
    ) -> None:
        """Memulai pemrosesan file dalam background daemon thread."""
        t = threading.Thread(
            target=self._process_queue,
            args=(module_name, function_name, file_list, output_folder),
            daemon=True,
        )
        t.start()

    def _log(self, message: str, level: str = "DEFAULT") -> None:
        """Mengirim pesan log ke antarmuka pengguna."""
        self.status_queue.put(("LOG", message, level))

    def _request_gui(self, task_name: str, **kwargs) -> Any:
        """Meminta input interaktif dari GUI (password atau konfirmasi timpa) dan menunggu respon."""
        result_q = queue.Queue()
        self.request_queue.put((task_name, kwargs, result_q))
        try:
            return result_q.get(timeout=300)
        except queue.Empty:
            return None

    def _authenticate_pdf(self, pdf_path: str, PdfReader: Any) -> Tuple[bool, Optional[str]]:
        """Memeriksa enkripsi PDF dan meminta password jika diperlukan.

        Returns (is_accessible, password)
        """
        if not PdfReader:
            return True, None

        reader = None
        try:
            reader = PdfReader(pdf_path)
            if not getattr(reader, "is_encrypted", False):
                return True, None

            # Coba decrypt dengan password kosong (owner-only encryption)
            try:
                if reader.decrypt(""):
                    return True, ""
            except Exception:
                pass

            if not getattr(reader, "is_encrypted", False):
                return True, ""

            base_name = os.path.basename(pdf_path)
            self._log(f"File '{base_name}' terproteksi.", level="INFO")
            pwd = self._request_gui(
                task_name="ask_password",
                text=f"Masukkan password:\n{base_name}",
                title="Password",
            )
            if pwd is None:
                self._log("Dibatalkan pengguna.", level="ERROR")
                return False, None

            try:
                res = reader.decrypt(pwd)
                if not res and getattr(reader, "is_encrypted", False):
                    self._log("Password salah.", level="ERROR")
                    return False, None
                return True, pwd
            except Exception:
                self._log("Password salah.", level="ERROR")
                return False, None

        except Exception as e:
            self._log(f"Error cek enkripsi: {e}", level="ERROR")
            return False, None
        finally:
            if reader:
                try:
                    reader.close()
                except Exception:
                    pass

    def _create_decrypted_copy(
        self,
        pdf_path: str,
        password: Optional[str],
        PdfReader: Any,
        PdfWriter: Any,
    ) -> Tuple[str, Optional[str]]:
        """Membuat salinan PDF yang sudah didekripsi jika diperlukan oleh parser.

        Returns (pdf_to_process, temp_pdf_path)
        """
        if password is None or not PdfWriter:
            if password is not None and not PdfWriter:
                self._log(
                    "pypdf tidak tersedia; tidak dapat membuat salinan didekripsi.",
                    level="INFO",
                )
            return pdf_path, None

        reader2 = None
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
            return tmp.name, tmp.name
        except Exception as e:
            self._log(f"Warning: gagal membuat salinan decrypted: {e}", level="ERROR")
            return pdf_path, None
        finally:
            if reader2:
                try:
                    reader2.close()
                except Exception:
                    pass

    def _check_overwrite_permission(self, excel_path: str) -> Optional[bool]:
        """Memeriksa apakah file excel tujuan sudah ada dan meminta konfirmasi user.

        Returns:
            True: izinkan overwrite
            False: lewati file ini
            None: batalkan seluruh proses
        """
        if not os.path.exists(excel_path):
            return True
        return self._request_gui(
            task_name="ask_overwrite",
            title="Overwrite?",
            message=f"File ada:\n{excel_path}\nTimpa?",
        )

    def _process_single_file(
        self,
        pdf_path: str,
        module_name: str,
        function_name: str,
        output_folder: str,
        PdfReader: Any,
        PdfWriter: Any,
    ) -> Optional[bool]:
        """Memproses satu file PDF dari awal sampai akhir via extract_statement."""
        is_ok, password = self._authenticate_pdf(pdf_path, PdfReader)
        if not is_ok:
            return False

        temp_pdf_path = None
        try:
            pdf_to_process, temp_pdf_path = self._create_decrypted_copy(
                pdf_path, password, PdfReader, PdfWriter
            )

            base_name = os.path.basename(pdf_path)
            file_name, _ = os.path.splitext(base_name)
            excel_path = os.path.join(output_folder, f"{file_name}.xlsx")

            overwrite_decision = self._check_overwrite_permission(excel_path)
            if overwrite_decision is None:
                return None  # Batalkan seluruh antrean
            if not overwrite_decision:
                return False  # Lewati file ini

            count = extract_statement(
                pdf_path=pdf_to_process,
                excel_path=excel_path,
                module_name=module_name,
                function_name=function_name,
                password=password,
                log_callback=self._log,
                progress_callback=lambda curr, tot, msg: self._log(msg, "INFO"),
            )

            if count > 0:
                self.status_queue.put(("FILE", excel_path, "SUCCESS"))
                return True
            else:
                return False

        finally:
            if temp_pdf_path and os.path.exists(temp_pdf_path):
                try:
                    os.remove(temp_pdf_path)
                except Exception:
                    pass

    def _process_queue(
        self,
        module_name: str,
        function_name: str,
        file_list: List[str],
        output_folder: str,
    ) -> None:
        """Loop pemrosesan antrean file PDF."""
        try:
            from pypdf import PdfReader, PdfWriter
        except ImportError:
            PdfReader = None
            PdfWriter = None

        files_processed = 0
        for pdf_path in file_list:
            pdf_path = pdf_path.strip()
            if not pdf_path:
                continue

            status = self._process_single_file(
                pdf_path, module_name, function_name, output_folder, PdfReader, PdfWriter
            )
            if status is None:
                break
            files_processed += 1

        if files_processed > 0:
            self._log("========================================", level="SEPARATOR")
            self._log("SEMUA PROSES SELESAI.\n", level="SUCCESS")
        else:
            self._log("Tidak ada file yang diproses.", level="INFO")

    def _run_module(
        self,
        module_name: str,
        function_name: str,
        pdf_path: str,
        excel_path: str,
        password: Optional[str] = None,
    ) -> Any:
        """Helper backward-compatibility untuk pemanggilan parser modul."""
        try:
            return _call_parser_func(
                module_name,
                function_name,
                pdf_path,
                excel_path,
                password,
                progress_callback=lambda curr, tot, msg: self._log(msg, "INFO"),
            )
        except Exception as e:
            return False


def is_empty_result(result: Any) -> bool:
    """Menentukan apakah hasil scraping berupa hasil kosong (0 baris/transaksi) tanpa error.

    Fallback aktif HANYA jika scraping selesai tanpa error tetapi menghasilkan 0 baris/transaksi.
    Jika result bernilai False, None, atau bertipe Exception, berarti terjadi error/kegagalan baca
    (misal: PDF rusak, password salah/dibatalkan, dsb), sehingga BUKAN empty_result dan fallback TIDAK aktif.
    """
    if result is False or result is None:
        return False
    if isinstance(result, (Exception, bool)):
        return False
    if isinstance(result, (int, float)):
        return result == 0
    if hasattr(result, "__len__"):
        return len(result) == 0
    return False


def _call_parser_func(
    module_name: str,
    function_name: str,
    pdf_path: str,
    excel_path: str,
    password: Optional[str] = None,
    progress_callback: Optional[Any] = None,
) -> Any:
    """Memanggil fungsi parser bank secara dinamis dengan parameter yang sesuai."""
    mod = importlib.import_module(module_name)
    func = getattr(mod, function_name, None)

    if not func and hasattr(mod, "parse"):
        func = getattr(mod, "parse")

    if not func:
        raise AttributeError(f"Fungsi '{function_name}' atau 'parse' tidak ditemukan pada modul {module_name}")

    sig = inspect.signature(func)
    call_args = {}
    if "pdf_path" in sig.parameters:
        call_args["pdf_path"] = pdf_path
    if "output_excel" in sig.parameters:
        call_args["output_excel"] = excel_path
    if "excel_path" in sig.parameters:
        call_args["excel_path"] = excel_path
    if "output_excel_path" in sig.parameters:
        call_args["output_excel_path"] = excel_path
    if "password" in sig.parameters and password:
        call_args["password"] = password
    if "progress_callback" in sig.parameters and progress_callback:
        call_args["progress_callback"] = progress_callback

    return func(**call_args)


def extract_statement(
    pdf_path: str,
    excel_path: str,
    module_name: str = "AUTO",
    function_name: str = "parse",
    password: Optional[str] = None,
    log_callback: Optional[Any] = None,
    progress_callback: Optional[Any] = None,
) -> int:
    """Titik masuk ekstraksi terpadu di lapisan inti (Scraping utama + Fallback OCR otomatis).

    Aturan alur:
    1. Selalu jalankan scraping terlebih dahulu.
    2. Jika scraping menghasilkan data (> 0), OCR TIDAK dijalankan dan engine OCR TIDAK dimuat (lazy).
    3. Jika scraping gagal karena error (PDF rusak, password salah/gagal, dsb), error dilaporkan
       apa adanya dan fallback OCR TIDAK dijalankan.
    4. Fallback OCR HANYA aktif jika scraping selesai membaca file TANPA error, tetapi hasilnya
       nol baris data / nol transaksi (is_empty_result).
    """
    def _log(msg: str, level: str = "DEFAULT"):
        if log_callback:
            log_callback(msg, level)

    if not os.path.exists(pdf_path):
        _log(f"File PDF tidak ditemukan: {pdf_path}", "ERROR")
        return 0

    is_scraping_empty = False

    # --- JALUR 1: AUTO DETECT ---
    if module_name == "AUTO":
        pdf_text = ""
        try:
            from pypdf import PdfReader
            reader = PdfReader(pdf_path)
            if reader.is_encrypted and password:
                reader.decrypt(password)
            if len(reader.pages) > 0:
                pdf_text = (reader.pages[0].extract_text() or "").upper()
            reader.close()
        except Exception as e:
            _log(f"Error membaca PDF ({type(e).__name__}): {e}", "ERROR")
            return 0

        prioritized = []
        others = []
        for b_mod, b_func, keywords in AUTO_BANKS:
            if any(kw.upper() in pdf_text for kw in keywords):
                prioritized.append((b_mod, b_func))
            else:
                others.append((b_mod, b_func))

        for b_mod, b_func in prioritized + others:
            try:
                _log(f"Auto-detect mencoba scraping: {b_mod}...", "INFO")
                res = _call_parser_func(b_mod, b_func, pdf_path, excel_path, password)
                if res and isinstance(res, int) and res > 0:
                    _log(f"Berhasil! Terdeteksi sebagai format {b_mod} ({res} transaksi).", "SUCCESS")
                    return res
            except Exception:
                continue

        # Jika seluruh parser scraping pada AUTO tidak menghasilkan data
        is_scraping_empty = True

    # --- JALUR 2: SPESIFIK BANK ---
    else:
        try:
            _log(f"Menjalankan ekstraksi scraping: {module_name}...", "RUN")
            scraping_res = _call_parser_func(
                module_name, function_name, pdf_path, excel_path, password
            )
        except Exception as e:
            # Error saat pembacaan file/PDF (PDF rusak, tidak valid, dsb)
            _log(f"Error saat scraping {module_name} ({type(e).__name__}): {e}", "ERROR")
            # Aturan: Error tidak ditutupi OCR
            return 0

        if scraping_res and isinstance(scraping_res, int) and scraping_res > 0:
            _log(f"Scraping {module_name} berhasil mengekstrak {scraping_res} transaksi.", "SUCCESS")
            # JANGAN muat modul OCR! Lazy import dipertahankan.
            return scraping_res

        if is_empty_result(scraping_res):
            is_scraping_empty = True
        else:
            _log(f"Scraping {module_name} gagal atau mengembalikan status tidak valid.", "ERROR")
            return 0

    # --- JALUR 3: FALLBACK OCR (HANYA JIKA is_empty_result TRUE) ---
    if not is_scraping_empty:
        return 0

    _log("Scraping tidak menghasilkan transaksi (0 data). Mengaktifkan fallback OCR otomatis...", "WARNING")

    # Cek dukungan bank
    supported_ocr_banks = ("parser.Livin", "parser.Mandiri", "AUTO", "Livin", "Mandiri")
    if module_name not in supported_ocr_banks:
        _log(
            f"Fallback OCR untuk format {module_name} belum tersedia [DATA BELUM TERSEDIA]. "
            "Saat ini hanya Mandiri/Livin yang didukung.",
            "WARNING"
        )
        return 0

    # Lazy import engine OCR
    try:
        from ocr_engine import is_ocr_available, validate_running_balance
        if not is_ocr_available():
            _log("Engine OCR (rapidocr-onnxruntime / pypdfium2) belum terpasang. Fallback OCR tidak dapat dijalankan.", "ERROR")
            return 0
    except ImportError as e:
        _log(f"Gagal memuat engine OCR: {e}", "ERROR")
        return 0

    # Eksekusi OCR
    try:
        _log("Menjalankan OCR lokal offline...", "INFO")
        from parser.Livin import parse_ocr as livin_parse_ocr
        ocr_res = livin_parse_ocr(
            pdf_path,
            output_excel_path=excel_path,
            password=password,
            progress_callback=progress_callback
        )

        if ocr_res and isinstance(ocr_res, int) and ocr_res > 0:
            _log(f"Fallback OCR berhasil mengekstrak {ocr_res} transaksi.", "SUCCESS")
            # Validasi saldo berjalan (logika pelaporan di UI/log, sel Excel tidak diubah)
            try:
                import pandas as pd
                df_check = pd.read_excel(excel_path)
                is_bal_valid, bal_warnings = validate_running_balance(df_check)
                if is_bal_valid:
                    _log("Validasi saldo: Rekonsiliasi matematis konsisten 100%.", "SUCCESS")
                else:
                    for w in bal_warnings[:3]:
                        _log(f"Peringatan saldo: {w}", "WARNING")
            except Exception as e_bal:
                _log(f"Peringatan validasi saldo: {e_bal}", "WARNING")

            return ocr_res
        else:
            _log("OCR selesai dijalankan, namun tidak ditemukan baris transaksi yang valid (kasus ganda kosong).", "WARNING")
            return 0

    except Exception as e_ocr:
        _log(f"Error saat eksekusi OCR ({type(e_ocr).__name__}): {e_ocr}", "ERROR")
        return 0