import os
import json
import pytest
import tempfile
from settings import Settings, SettingsManager, SCHEMA_VERSION


class TestSettingsModelAndManager:
    """Pengujian komprehensif untuk Settings dan SettingsManager."""

    def test_default_values_match_baseline(self):
        """Nilai default harus persis sama dengan perilaku sistem baseline."""
        s = Settings()
        assert s.schema_version == SCHEMA_VERSION
        assert s.appearance_mode == "System"
        assert s.theme_name == "Default (Blue)"
        assert s.thousands_separator == ","
        assert s.decimal_separator == "."
        assert s.decimal_places == 2
        assert s.negative_format == "-1234"
        assert s.output_number_type == "text"
        assert s.date_format == "DD/MM/YYYY"
        assert s.output_folder_mode == "last_used"
        assert s.overwrite_mode == "ask"
        assert s.auto_open_output == "none"
        assert s.ocr_fallback_enabled is True
        assert s.mask_sensitive_logs is True
        assert s.auto_cleanup_temp_files is True

        is_valid, err = s.validate()
        assert is_valid is True
        assert err is None

    def test_cross_validation_same_separators(self):
        """Pemisah ribuan dan desimal tidak boleh sama."""
        s = Settings(thousands_separator=".", decimal_separator=".")
        is_valid, err = s.validate()
        assert is_valid is False
        assert "tidak boleh menggunakan karakter yang sama" in err

        s2 = Settings(thousands_separator=",", decimal_separator=",")
        is_valid, err = s2.validate()
        assert is_valid is False

        # Pemisah ribuan kosong diperbolehkan jika desimal . atau ,
        s3 = Settings(thousands_separator="", decimal_separator=".")
        is_valid, err = s3.validate()
        assert is_valid is True

    def test_invalid_option_values(self):
        """Validasi nilai di luar domain yang diizinkan."""
        assert Settings(decimal_places=-1).validate()[0] is False
        assert Settings(decimal_places=10).validate()[0] is False
        assert Settings(output_number_type="binary").validate()[0] is False
        assert Settings(negative_format="red_text").validate()[0] is False
        assert Settings(appearance_mode="Neon").validate()[0] is False

    def test_save_and_load_roundtrip(self, tmp_path):
        """Penyimpanan atomik dan pembacaan ulang harus konsisten."""
        cfg_file = str(tmp_path / "settings_test.json")
        mgr = SettingsManager(cfg_file)

        custom = Settings(
            appearance_mode="Dark",
            theme_name="Special Edition (Pink)",
            thousands_separator=".",
            decimal_separator=",",
            decimal_places=2,
            negative_format="(1234)",
            output_number_type="numeric",
            date_format="YYYY-MM-DD",
            ocr_fallback_enabled=False
        )

        success = mgr.save(custom)
        assert success is True
        assert os.path.exists(cfg_file)

        loaded = mgr.load()
        assert loaded.appearance_mode == "Dark"
        assert loaded.theme_name == "Special Edition (Pink)"
        assert loaded.thousands_separator == "."
        assert loaded.decimal_separator == ","
        assert loaded.output_number_type == "numeric"
        assert loaded.ocr_fallback_enabled is False

    def test_corrupted_file_handling(self, tmp_path):
        """File JSON yang korup atau tidak valid harus fallback ke default dan membuat backup .bak."""
        cfg_file = str(tmp_path / "corrupted_settings.json")
        with open(cfg_file, "w") as f:
            f.write("{this is not valid json!]")

        mgr = SettingsManager(cfg_file)
        loaded = mgr.load()

        # Harus menghasilkan default tanpa crash
        assert isinstance(loaded, Settings)
        assert loaded.thousands_separator == ","
        assert loaded.decimal_separator == "."

        # Berkas .bak harus terbentuk
        bak_file = f"{cfg_file}.bak"
        assert os.path.exists(bak_file)

    def test_schema_migration_from_v0(self, tmp_path):
        """Migrasi berkas lama (yang hanya punya appearance_mode dan theme_name)."""
        cfg_file = str(tmp_path / "legacy_settings.json")
        with open(cfg_file, "w") as f:
            json.dump({"appearance_mode": "Light", "theme_name": "Special Edition (Pink)"}, f)

        mgr = SettingsManager(cfg_file)
        loaded = mgr.load()

        assert loaded.schema_version == SCHEMA_VERSION
        assert loaded.appearance_mode == "Light"
        assert loaded.theme_name == "Special Edition (Pink)"
        # Nilai baru otomatis terisi default
        assert loaded.thousands_separator == ","
        assert loaded.decimal_separator == "."

    def test_export_and_import(self, tmp_path):
        """Fungsi ekspor dan impor berkas JSON."""
        cfg_file = str(tmp_path / "orig.json")
        export_file = str(tmp_path / "exported.json")
        mgr = SettingsManager(cfg_file)

        settings = Settings(thousands_separator=" ", decimal_separator=",")
        mgr.save(settings)

        assert mgr.export_settings(export_file, settings) is True
        assert os.path.exists(export_file)

        new_cfg_file = str(tmp_path / "imported.json")
        new_mgr = SettingsManager(new_cfg_file)
        imported = new_mgr.import_settings(export_file)

        assert imported is not None
        assert imported.thousands_separator == " "
        assert imported.decimal_separator == ","

    def test_ocr_fallback_toggle_disabled(self, tmp_path):
        """Saat ocr_fallback_enabled False, OCR tidak dipanggil meski scraping menghasilkan 0 data."""
        from unittest.mock import patch
        from logic import extract_statement

        fake_pdf = str(tmp_path / "empty.pdf")
        with open(fake_pdf, "wb") as f:
            f.write(b"%PDF-1.4 dummy")
        out_excel = str(tmp_path / "out.xlsx")

        settings_no_ocr = Settings(ocr_fallback_enabled=False)

        with patch("parser.Livin.parse", return_value=0) as mock_scrape:
            with patch("parser.Livin.parse_ocr") as mock_ocr:
                logs = []
                count = extract_statement(
                    fake_pdf,
                    out_excel,
                    module_name="parser.Livin",
                    log_callback=lambda m, l: logs.append((m, l)),
                    settings=settings_no_ocr
                )
                mock_scrape.assert_called_once()
                mock_ocr.assert_not_called()
                assert count == 0
                assert any("dinonaktifkan dalam Pengaturan" in m for m, l in logs)

    def test_mask_sensitive_logs_in_logic(self):
        """CoreLogic._log menyamarkan nomor rekening dan saldo jika mask_sensitive_logs True."""
        import queue
        from logic import CoreLogic

        status_q = queue.Queue()
        request_q = queue.Queue()
        core = CoreLogic(status_q, request_q)
        core._active_settings = Settings(mask_sensitive_logs=True)

        core._log("Rekening 1234567890 memiliki saldo Rp 50.000.000,00.")
        msg_type, content, level = status_q.get()
        assert "1234567890" not in content
        assert "50.000.000,00" not in content
        assert "1234**7890" in content

        # Saat dimatikan, log ditampilkan utuh
        core._active_settings = Settings(mask_sensitive_logs=False)
        core._log("Rekening 1234567890 memiliki saldo Rp 50.000.000,00.")
        msg_type, content, level = status_q.get()
        assert "1234567890" in content
        assert "50.000.000,00" in content

    def test_overwrite_mode_rename(self, tmp_path):
        """Saat overwrite_mode 'rename', file yang sudah ada diberi nomor urut otomatis."""
        from logic import CoreLogic
        import queue

        core = CoreLogic(queue.Queue(), queue.Queue())
        out_dir = str(tmp_path)
        existing = os.path.join(out_dir, "test.xlsx")
        with open(existing, "w") as f:
            f.write("existing content")

        fake_pdf = os.path.join(out_dir, "test.pdf")
        with open(fake_pdf, "wb") as f:
            f.write(b"%PDF-1.4 dummy")

        settings_rename = Settings(overwrite_mode="rename", file_naming_pattern="{nama_pdf}")

        # Jalankan _process_single_file dengan mock extract_statement
        from unittest.mock import patch
        with patch("logic.extract_statement", return_value=10) as mock_extract:
            core._process_single_file(
                fake_pdf,
                module_name="parser.Mandiri",
                function_name="parse",
                output_folder=out_dir,
                PdfReader=None,
                PdfWriter=None,
                settings=settings_rename
            )
            # Pastikan dipanggil dengan excel_path yang direname ke test (1).xlsx
            call_kwargs = mock_extract.call_args[1]
            assert "test (1).xlsx" in call_kwargs["excel_path"]
            assert os.path.exists(existing)

    def test_settings_dialog_headless_interaction(self, tmp_path):
        """Test interaksi SettingsDialog: ganti kategori, preview langsung, simpan."""
        import customtkinter as ctk
        from gui import FinextractUI, SettingsDialog

        cfg_file = str(tmp_path / "gui_settings.json")
        mgr = SettingsManager(cfg_file)

        root = FinextractUI(settings_manager=mgr)
        root.withdraw()

        dlg = SettingsDialog(root, settings_manager=mgr)
        dlg.withdraw()

        # Ubah kategori
        dlg._select_category("date")
        dlg.var_date_format.set("YYYY-MM-DD (Contoh: 2026-10-17)")
        dlg._update_live_preview()
        assert "2026-10-17" in dlg.lbl_preview_text.cget("text")

        # Test validasi silang (pemisah sama)
        dlg.var_thousands.set("Titik (.) [Standar ID/EU]")
        dlg.var_decimal.set("Titik (.) [Standar US/UK]")
        dlg._update_live_preview()
        assert "tidak boleh" in dlg.lbl_validation_err.cget("text")
        assert dlg.btn_save.cget("state") == "disabled"

        # Kembalikan ke yang valid
        dlg.var_decimal.set("Koma (,) [Standar ID/EU]")
        dlg._update_live_preview()
        assert dlg.lbl_validation_err.cget("text") == ""
        assert dlg.btn_save.cget("state") == "normal"

        # Simpan
        dlg.on_save()
        loaded = mgr.load()
        assert loaded.thousands_separator == "."
        assert loaded.decimal_separator == ","
        assert loaded.date_format == "YYYY-MM-DD"

        dlg.destroy()
        root.destroy()

