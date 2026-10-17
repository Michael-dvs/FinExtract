import os
import shutil
import tempfile
import hashlib
import queue
import pytest
import pandas as pd

from config import AUTO_BANKS, THEME_CONFIG
from logic import CoreLogic
from parser.BRI import extract_bri_text
from parser.Mandiri import process_bank_statement
from parser.Livin import (
    process_livin_statement,
    clean_text,
    format_date_excel,
    is_date_start,
    is_money,
    align_bank_row,
)

BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
REKKOR_DIR = os.path.join(BASE_DIR, "Rekkor")


class TestCharacterizationParsers:
    """Characterization tests for the 3 target banks (BRI, Mandiri, Livin)

    These tests lock current behavior and verify exact output structures
    and content hashes without exposing sensitive financial records.
    """

    def test_bri_characterization(self):
        pdf_path = os.path.join(REKKOR_DIR, "BRI.pdf")
        assert os.path.exists(pdf_path), "File Rekkor/BRI.pdf must exist"

        with tempfile.TemporaryDirectory() as tmpdir:
            out_excel = os.path.join(tmpdir, "bri_output.xlsx")
            row_count = extract_bri_text(pdf_path, out_excel)

            assert row_count == 125
            assert os.path.exists(out_excel)

            df = pd.read_excel(out_excel)
            assert df.shape == (125, 6)
            expected_cols = ["transaction_date", "description", "user_id", "debit", "credit", "balance"]
            assert list(df.columns) == expected_cols

            df_hash = hashlib.sha256(pd.util.hash_pandas_object(df).values).hexdigest()
            assert df_hash == "b90c1ca664e7bf366469a7bdfb15f1bc8ef5882906c0f160116042c21ef42162"

    def test_mandiri_characterization(self):
        pdf_path = os.path.join(REKKOR_DIR, "Mandiri.pdf")
        assert os.path.exists(pdf_path), "File Rekkor/Mandiri.pdf must exist"

        with tempfile.TemporaryDirectory() as tmpdir:
            out_excel = os.path.join(tmpdir, "mandiri_output.xlsx")
            row_count = process_bank_statement(pdf_path, out_excel)

            assert row_count == 190
            assert os.path.exists(out_excel)

            df = pd.read_excel(out_excel)
            assert df.shape == (190, 6)
            expected_cols = ["Posting Date", "Remark", "Reference No", "Debit", "Credit", "Balance"]
            assert list(df.columns) == expected_cols

            df_hash = hashlib.sha256(pd.util.hash_pandas_object(df).values).hexdigest()
            assert df_hash == "0db48064990af03bf9945116fd94235dea764a744819ca32fd8f3a2981837198"

    def test_livin_characterization(self):
        pdf_path = os.path.join(REKKOR_DIR, "Livin.pdf")
        assert os.path.exists(pdf_path), "File Rekkor/Livin.pdf must exist"

        with tempfile.TemporaryDirectory() as tmpdir:
            # Copy to tmpdir to prevent writing into Rekkor folder
            tmp_pdf = os.path.join(tmpdir, "Livin.pdf")
            shutil.copyfile(pdf_path, tmp_pdf)

            row_count = process_livin_statement(tmp_pdf)
            expected_excel = os.path.join(tmpdir, "Livin_livin.xlsx")

            assert row_count == 141
            assert os.path.exists(expected_excel)

            df = pd.read_excel(expected_excel)
            assert df.shape == (141, 6)
            expected_cols = ["Posting Date", "Remark", "Reference No", "Debit", "Credit", "Balance"]
            assert list(df.columns) == expected_cols

            df_hash = hashlib.sha256(pd.util.hash_pandas_object(df).values).hexdigest()
            assert df_hash == "ff3950b8230b0b33be54ddd7d894b1ac05c12242f56d2b07b7e495ce9717ef52"

    def test_standardized_parse_interface(self):
        """Verify that all parsers expose the standardized parse() interface."""
        import parser.BRI as mod_bri
        import parser.Mandiri as mod_man
        import parser.Livin as mod_liv
        import parser.BNI as mod_bni
        import parser.OCBC as mod_ocbc

        for mod in [mod_bri, mod_man, mod_liv, mod_bni, mod_ocbc]:
            assert hasattr(mod, "parse"), f"Module {mod.__name__} must expose parse()"
            assert callable(mod.parse)

        with tempfile.TemporaryDirectory() as tmpdir:
            # Test Livin with explicit output_excel_path
            out_livin = os.path.join(tmpdir, "custom_livin.xlsx")
            res = mod_liv.parse(os.path.join(REKKOR_DIR, "Livin.pdf"), output_excel_path=out_livin)
            assert res == 141
            assert os.path.exists(out_livin)
            df = pd.read_excel(out_livin)
            assert df.shape == (141, 6)


class TestHelperFunctions:
    """Unit tests for helper routines used in parsers and logic."""

    def test_livin_clean_text(self):
        assert clean_text(None) == ""
        assert clean_text("  hello \n world  ") == "hello world"

    def test_livin_format_date_excel(self):
        assert format_date_excel("") == ""
        assert format_date_excel("01 Dec 2025, 10:00") == "01/12/2025"
        assert format_date_excel("invalid_date") == "invalid_date"

    def test_livin_is_date_start(self):
        assert is_date_start("01 Dec 2025") is True
        assert is_date_start("15 Jan 2024 Transfer") is True
        assert is_date_start("Saldo Awal") is False
        assert is_date_start(None) is False

    def test_livin_is_money(self):
        assert is_money("10,000.00") is True
        assert is_money("-500.00") is True
        assert is_money("abc") is False
        assert is_money("") is False

    def test_livin_align_bank_row_empty(self):
        result = align_bank_row([])
        assert result == [""] * 6

    def test_livin_align_bank_row_insufficient(self):
        result = align_bank_row(["01 Dec 2025", "desc"])
        assert result == ["01 Dec 2025", "PARSE ERROR", "", "0", "0", "0"]

    def test_swap_columns(self):
        from utils import swap_columns
        df = pd.DataFrame({"Debit": [100, 200], "Credit": [0, 50]})
        swap_columns(df, "Debit", "Credit")
        assert list(df["Debit"]) == [0, 50]
        assert list(df["Credit"]) == [100, 200]

    def test_open_system_path_invalid(self):
        from utils import open_system_path
        assert open_system_path("") is False
        assert open_system_path("/non/existent/path/here") is False


class TestCoreLogicSettings:
    """Tests for CoreLogic settings persistence without touching user home config."""

    def test_settings_load_default(self, monkeypatch, tmp_path):
        dummy_file = str(tmp_path / "non_existent.json")
        logic = CoreLogic(queue.Queue(), queue.Queue())
        monkeypatch.setattr(logic, "config_file", dummy_file)

        settings = logic.load_settings()
        assert settings == {"appearance_mode": "System", "theme_name": "Default (Blue)"}

    def test_settings_save_and_load(self, monkeypatch, tmp_path):
        test_file = str(tmp_path / "test_settings.json")
        logic = CoreLogic(queue.Queue(), queue.Queue())
        monkeypatch.setattr(logic, "config_file", test_file)

        to_save = {"appearance_mode": "Dark", "theme_name": "Special Edition (Pink)"}
        logic.save_settings(to_save)
        loaded = logic.load_settings()
        assert loaded == to_save

    def test_check_overwrite_permission_nonexistent(self, tmp_path):
        logic = CoreLogic(queue.Queue(), queue.Queue())
        non_existent_file = str(tmp_path / "new_file.xlsx")
        assert logic._check_overwrite_permission(non_existent_file) is True

    def test_create_decrypted_copy_no_password(self):
        logic = CoreLogic(queue.Queue(), queue.Queue())
        orig_path = "/dummy/sample.pdf"
        res_path, tmp_path = logic._create_decrypted_copy(orig_path, None, None, None)
        assert res_path == orig_path
        assert tmp_path is None
