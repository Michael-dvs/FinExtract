import os
import json
import threading
import importlib
import traceback
import tempfile
import queue
from config import AUTO_BANKS

try:
    from pypdf import PdfReader, PdfWriter
except ImportError:
    PdfReader = None
    PdfWriter = None

class CoreLogic:
    def __init__(self, status_queue, request_queue):
        self.status_queue = status_queue
        self.request_queue = request_queue
        self.config_file = os.path.join(os.path.expanduser("~"), "FinExtract_Settings.json")
        self.AUTO_BANKS = AUTO_BANKS

    def load_settings(self):
        default = {"appearance_mode": "System", "theme_name": "Default (Blue)"}
        try:
            if os.path.exists(self.config_file):
                with open(self.config_file, "r") as f:
                    return json.load(f)
        except Exception: pass
        return default

    def save_settings(self, settings):
        try:
            with open(self.config_file, "w") as f:
                json.dump(settings, f)
        except Exception: pass

    def start_processing_thread(self, module_name, function_name, file_list, output_folder):
        t = threading.Thread(
            target=self._process_queue, 
            args=(module_name, function_name, file_list, output_folder), 
            daemon=True
        )
        t.start()

    def _log(self, message, level="DEFAULT"):
        self.status_queue.put(("LOG", message, level))

    def _request_gui(self, task_name, **kwargs):
        """Meminta GUI melakukan sesuatu (misal popup password) dan menunggu hasil"""
        result_q = queue.Queue()
        self.request_queue.put((task_name, kwargs, result_q))
        try:
            return result_q.get(timeout=300)
        except queue.Empty:
            return None    

    def _process_queue(self, module_name, function_name, file_list, output_folder):
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
                            self._log(f"File '{os.path.basename(pdf_path)}' terproteksi.", level="INFO")
                            password = self._request_gui(task_name='ask_password', text=f"Masukkan password:\n{os.path.basename(pdf_path)}", title="Password")
                            if password is None:
                                self._log("Dibatalkan pengguna.", level="ERROR")
                                continue

                            # Coba decrypt dengan password yg diberikan pengguna
                            try:
                                res = reader.decrypt(password) if reader else False
                                # Jika decrypt gagal, laporkan kesalahan dan lanjut ke file berikutnya
                                if not res and getattr(reader, "is_encrypted", False):
                                    self._log("Password salah.", level="ERROR")
                                    continue
                            except Exception:
                                self._log("Password salah.", level="ERROR")
                                continue
                    except Exception as e:
                        self._log(f"Error saat mencoba decrypt: {e}", level="ERROR")
                        continue
            except Exception as e:
                self._log(f"Error cek enkripsi: {e}", level="ERROR")
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
                    self._log(f"Warning: gagal membuat salinan decrypted: {e}", level="ERROR")
            elif password is not None and not PdfWriter:
                self._log("pypdf tidak tersedia; tidak dapat membuat salinan didekripsi. Beberapa file terenkripsi mungkin gagal dibuka oleh pdfplumber.", level="INFO")

            base_name = os.path.basename(pdf_path)
            file_name, _ = os.path.splitext(base_name)
            excel_path = os.path.join(output_folder, f"{file_name}.xlsx")
            if os.path.exists(excel_path):
                response = self._request_gui(task_name='ask_overwrite', title="Overwrite?", message=f"File ada:\n{excel_path}\nTimpa?")
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
                        self._log(f"Mencoba format: {b_mod}...", level="INFO")
                        result = self._run_module(b_mod, b_func, pdf_to_process, excel_path, password)
                        if result and isinstance(result, int) and result > 0:
                            self._log(f"Berhasil! Terdeteksi sebagai format {b_mod}.", level="SUCCESS")
                            self.status_queue.put(("FILE", excel_path, "SUCCESS"))
                            found = True
                            break
                    except Exception:
                        continue
                if not found:
                    self._log(f"Gagal mendeteksi format untuk: {base_name}", level="ERROR")
            else:
                result = self._run_module(module_name, function_name, pdf_to_process, excel_path, password)
                if result and isinstance(result, int) and result > 0:
                    self.status_queue.put(("FILE", excel_path, "SUCCESS"))

            if temp_pdf_path and os.path.exists(temp_pdf_path):
                try: os.remove(temp_pdf_path)
                except Exception: pass
            files_processed += 1

        if files_processed > 0:
            self._log("========================================", level="SEPARATOR")
            self._log("SEMUA PROSES SELESAI.\n", level="SUCCESS")
        else:
            self._log("Tidak ada file yang diproses.", level="INFO")   

    def _run_module(self, module_name, function_name, pdf_path, excel_path, password=None):
        try:
            mod = importlib.import_module(module_name)
            importlib.reload(mod)
            func = getattr(mod, function_name, None)
            
            if not func:
                self._log(f"Fungsi {function_name} tidak ditemukan.", "ERROR")
                return False

            self._log("Running extraction...", "RUN")
            # Inspeksi argumen function
            import inspect
            sig = inspect.signature(func)
            call_args = {}
            if 'pdf_path' in sig.parameters: call_args['pdf_path'] = pdf_path
            if 'output_excel' in sig.parameters: call_args['output_excel'] = excel_path
            if 'excel_path' in sig.parameters: call_args['excel_path'] = excel_path
            if 'output_excel_path' in sig.parameters: call_args['output_excel_path'] = excel_path
            if 'password' in sig.parameters and password: call_args['password'] = password

            return func(**call_args)
            
        except Exception as e:
            self._log(f"Error: {e}", "ERROR")
            return False
        
        