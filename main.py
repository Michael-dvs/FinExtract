import sys
import os
import queue
import customtkinter as ctk
from tkinter import messagebox

# Import modul lokal
from gui import FinextractUI
from logic import CoreLogic

class MainController:
    def __init__(self):
        #Setup Queues untuk komunikasi Thread
        self.status_queue = queue.Queue()
        self.request_queue = queue.Queue()

        #Inisialisasi Logic
        self.logic = CoreLogic(self.status_queue, self.request_queue)
        
        #Inisialisasi GUI
        self.ui = FinextractUI(settings_callback=self.save_settings)
        
        #Load Settings & Apply
        self.settings = self.logic.load_settings()
        ctk.set_appearance_mode(self.settings.get("appearance_mode", "System"))
        self.ui.apply_theme(self.settings.get("theme_name", "Default (Blue)"))

        #Hubungkan Action GUI ke Logic
        self.ui.set_process_callback(self.start_processing)

        #Mulai Loop Pengecekan Queue
        self.ui.after(100, self.check_queues)

    def save_settings(self):
        """Callback saat user mengubah setting di GUI"""
        new_settings = {
            "appearance_mode": ctk.get_appearance_mode(),
            "theme_name": self.ui.current_theme_key
        }
        self.logic.save_settings(new_settings)

    def start_processing(self, module_name, function_name):
        """Dipanggil saat tombol bank diklik"""
        input_files = self.ui.get_input_files()
        output_folder = self.ui.get_output_folder()

        if not input_files or not input_files[0]:
            messagebox.showerror("Error", "Harap pilih file input PDF.")
            return
        if not output_folder:
            messagebox.showerror("Error", "Harap pilih folder output.")
            return

        self.ui.disable_open_buttons()
        self.logic.start_processing_thread(module_name, function_name, input_files, output_folder)

    def check_queues(self):
        """Loop utama untuk update UI dari background thread"""
        try:
            #Cek Log
            while True:
                msg_type, msg_content, level = self.status_queue.get_nowait()
                if msg_type == "LOG":
                    self.ui.log_message(msg_content, level)
                elif msg_type == "FILE":
                    self.ui.enable_open_buttons(msg_content)
        except queue.Empty:
            pass

        try:
            #Cek Request (Popup Password / Overwrite)
            while True:
                task_name, kwargs, result_q = self.request_queue.get_nowait()
                
                if task_name == 'ask_password':
                    result = self.ui.ask_password(kwargs.get('title'), kwargs.get('text'))
                    result_q.put(result)
                
                elif task_name == 'ask_overwrite':
                    result = self.ui.ask_overwrite(kwargs.get('title'), kwargs.get('message'))
                    result_q.put(result)
                    
        except queue.Empty:
            pass
        
        # Jadwalkan ulang
        self.ui.after(100, self.check_queues)

    def run(self):
        self.ui.mainloop()

if __name__ == '__main__':
    base_dir = os.path.dirname(os.path.abspath(__file__))
    if base_dir not in sys.path:
        sys.path.insert(0, base_dir)

    app = MainController()
    app.run()
