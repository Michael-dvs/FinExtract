import os
import sys
import shutil
import tempfile
import hashlib
from unittest.mock import patch, MagicMock
import pytest
import pandas as pd
import openpyxl

from logic import extract_statement, is_empty_result
from tests.excel_comparator import compare_excel_workbooks

BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
REKKOR_DIR = os.path.join(BASE_DIR, "Rekkor")
SCRATCH_DIR = "/Users/michaelarhdyn/.gemini/antigravity-ide/brain/fa8c9aa9-d258-43c7-a53d-3295ddcb8831/scratch"
if not os.path.exists(SCRATCH_DIR):
    SCRATCH_DIR = os.path.join(BASE_DIR, "scratch")


class TestOCRExtractionPipeline:
    """Rangkaian pengujian titik masuk terpadu ekstraksi mutasi rekening (Fase 3).

    Semua pengujian berjalan melalui titik masuk `extract_statement` yang sama
    dengan yang digunakan oleh antarmuka GUI.
    """

    def test_is_empty_result_definition(self):
        """Memverifikasi definisi konkret dari hasil scraping kosong (is_empty_result)."""
        assert is_empty_result(0) is True
        assert is_empty_result(0.0) is True
        assert is_empty_result([]) is True
        assert is_empty_result(pd.DataFrame()) is True

        # Non-empty results
        assert is_empty_result(1) is False
        assert is_empty_result(141) is False
        assert is_empty_result(["row1"]) is False

        # Error / failure states MUST NOT be considered empty results
        assert is_empty_result(False) is False
        assert is_empty_result(None) is False
        assert is_empty_result(Exception("PDF Corrupt")) is False

    def test_lazy_loading_and_fallback_not_falsely_triggered(self):
        """Poin 4: Untuk PDF berbasis teks, OCR tidak dipanggil dan engine OCR tidak dimuat (lazy)."""
        pdf_path = os.path.join(REKKOR_DIR, "Livin.pdf")
        if not os.path.exists(pdf_path):
            pytest.skip("Rekkor/Livin.pdf tidak ditemukan.")

        with tempfile.TemporaryDirectory() as tmpdir:
            out_excel = os.path.join(tmpdir, "lazy_test.xlsx")

            # Pastikan modul OCR belum dimuat
            for mod in list(sys.modules.keys()):
                if "ocr_engine" in mod or "rapidocr" in mod:
                    del sys.modules[mod]

            with patch("parser.Livin.parse_ocr") as mock_ocr:
                count = extract_statement(pdf_path, out_excel, module_name="parser.Livin")

                # Membuktikan fungsi OCR tidak dipanggil
                mock_ocr.assert_not_called()
                assert count == 141

                # Membuktikan engine OCR tidak pernah dimuat ke dalam sys.modules
                assert "ocr_engine" not in sys.modules, "Engine OCR termuat padahal scraping berhasil!"
                assert "rapidocr_onnxruntime" not in sys.modules

    def test_fallback_triggered_precisely_once_on_empty(self):
        """Poin 5: Jika scraping menghasilkan 0 baris tanpa error, OCR dipanggil tepat 1 kali."""
        with tempfile.TemporaryDirectory() as tmpdir:
            fake_pdf = os.path.join(tmpdir, "empty_table.pdf")
            with open(fake_pdf, "wb") as f:
                f.write(b"%PDF-1.4 dummy")
            out_excel = os.path.join(tmpdir, "fallback_test.xlsx")

            with patch("parser.Livin.parse", return_value=0) as mock_scrape:
                with patch("parser.Livin.parse_ocr", return_value=25) as mock_ocr:
                    logs = []
                    count = extract_statement(
                        fake_pdf,
                        out_excel,
                        module_name="parser.Livin",
                        log_callback=lambda m, l: logs.append((m, l))
                    )

                    mock_scrape.assert_called_once()
                    mock_ocr.assert_called_once()
                    assert count == 25
                    assert any("Mengaktifkan fallback OCR" in m for m, l in logs)

    def test_error_not_masked_by_ocr(self):
        """Poin 6: Error membaca PDF (rusak/invalid) tidak ditutupi oleh OCR dan dilaporkan apa adanya."""
        with tempfile.TemporaryDirectory() as tmpdir:
            corrupt_pdf = os.path.join(tmpdir, "corrupted.pdf")
            with open(corrupt_pdf, "wb") as f:
                f.write(b"%PDF-corrupted-binary-header-trash")
            out_excel = os.path.join(tmpdir, "error_test.xlsx")

            logs = []
            with patch("parser.Livin.parse_ocr") as mock_ocr:
                count = extract_statement(
                    corrupt_pdf,
                    out_excel,
                    module_name="parser.Livin",
                    log_callback=lambda m, l: logs.append((m, l))
                )

                # Fallback OCR TIDAK boleh dijalankan ketika scraping gagal karena error
                mock_ocr.assert_not_called()
                assert count == 0
                assert not os.path.exists(out_excel)
                assert any(l == "ERROR" for m, l in logs)

    def test_caseless_double_empty(self):
        """Poin 7: Kasus ganda kosong (scraping 0 dan OCR 0) tertangani bersih tanpa crash."""
        with tempfile.TemporaryDirectory() as tmpdir:
            fake_pdf = os.path.join(tmpdir, "no_data.pdf")
            with open(fake_pdf, "wb") as f:
                f.write(b"%PDF-1.4 dummy")
            out_excel = os.path.join(tmpdir, "double_empty.xlsx")

            with patch("parser.Livin.parse", return_value=0):
                with patch("parser.Livin.parse_ocr", return_value=0):
                    logs = []
                    count = extract_statement(
                        fake_pdf,
                        out_excel,
                        module_name="parser.Livin",
                        log_callback=lambda m, l: logs.append((m, l))
                    )
                    assert count == 0
                    assert any("kasus ganda kosong" in m for m, l in logs)

    def test_no_scraping_regression_livin_and_mandiri(self):
        """Poin 3: Tanpa regresi scraping pada mandiri.pdf dan livin.pdf."""
        livin_path = os.path.join(REKKOR_DIR, "Livin.pdf")
        mandiri_path = os.path.join(REKKOR_DIR, "Mandiri.pdf")

        with tempfile.TemporaryDirectory() as tmpdir:
            if os.path.exists(livin_path):
                out_liv = os.path.join(tmpdir, "livin_baseline.xlsx")
                count_liv = extract_statement(livin_path, out_liv, module_name="parser.Livin")
                assert count_liv == 141
                df_liv = pd.read_excel(out_liv)
                hash_liv = hashlib.sha256(pd.util.hash_pandas_object(df_liv).values).hexdigest()
                assert hash_liv == "ff3950b8230b0b33be54ddd7d894b1ac05c12242f56d2b07b7e495ce9717ef52"

            if os.path.exists(mandiri_path):
                out_man = os.path.join(tmpdir, "mandiri_baseline.xlsx")
                count_man = extract_statement(mandiri_path, out_man, module_name="parser.Mandiri")
                assert count_man == 190
                df_man = pd.read_excel(out_man)
                hash_man = hashlib.sha256(pd.util.hash_pandas_object(df_man).values).hexdigest()
                assert hash_man == "0db48064990af03bf9945116fd94235dea764a744819ca32fd8f3a2981837198"

    def test_unified_entry_point_fallback_on_real_scan(self):
        """Poin 1: Ekstraksi scan mandiri_ocr / livin_ocr jatuh otomatis ke OCR dan tervalidasi saldo 100%."""
        pdf_path = os.path.join(REKKOR_DIR, "Livin_ocr.pdf")
        if not os.path.exists(pdf_path):
            pytest.skip("Rekkor/Livin_ocr.pdf tidak ditemukan.")

        with tempfile.TemporaryDirectory() as tmpdir:
            out_excel = os.path.join(tmpdir, "auto_fallback_result.xlsx")
            logs = []
            count = extract_statement(
                pdf_path,
                out_excel,
                module_name="parser.Livin",
                log_callback=lambda m, l: logs.append((m, l))
            )

            assert count == 93, f"Jumlah baris diharapkan 93, didapatkan {count}"
            assert os.path.exists(out_excel)
            assert any("Mengaktifkan fallback OCR otomatis" in m for m, l in logs)
            assert any("Validasi saldo: Rekonsiliasi matematis konsisten 100%." in m for m, l in logs)

    def test_mandiri_ocr_file_availability(self):
        """Memeriksa keberadaan mandiri_ocr.pdf sesuai kriteria uji kesetaraan."""
        mandiri_ocr_path = os.path.join(REKKOR_DIR, "mandiri_ocr.pdf")
        if not os.path.exists(mandiri_ocr_path):
            pytest.skip("[DATA BELUM TERSEDIA] Berkas Rekkor/mandiri_ocr.pdf tidak ditemukan di workspace.")

    def test_ocr_determinism_3_runs(self):
        """Poin 2: Determinisme - Ekstraksi OCR wajib stabil 100% identik pada 3 kali eksekusi berturut-turut."""
        pdf_path = os.path.join(REKKOR_DIR, "Livin_ocr.pdf")
        if not os.path.exists(pdf_path):
            pytest.skip("Rekkor/Livin_ocr.pdf tidak ditemukan.")

        hashes = []
        with tempfile.TemporaryDirectory() as tmpdir:
            for i in range(3):
                out_excel = os.path.join(tmpdir, f"run_{i+1}.xlsx")
                count = extract_statement(pdf_path, out_excel, module_name="parser.Livin")
                assert count == 93

                df = pd.read_excel(out_excel)
                df_hash = hashlib.sha256(pd.util.hash_pandas_object(df).values).hexdigest()
                hashes.append(df_hash)

            assert len(set(hashes)) == 1, f"Output OCR tidak deterministik: {hashes}"

    def test_comparator_negative_test(self):
        """Poin 8: Negative test pembanding - Pastikan pembanding melaporkan perbedaan jika 1 sel diubah."""
        ref_path = os.path.join(SCRATCH_DIR, "ref_Livin.xlsx")
        if not os.path.exists(ref_path):
            ref_path = os.path.join(REKKOR_DIR, "Mandiri.xlsx")
        if not os.path.exists(ref_path):
            pytest.skip("Berkas referensi Excel tidak ditemukan untuk negative test.")

        with tempfile.TemporaryDirectory() as tmpdir:
            tampered_path = os.path.join(tmpdir, "tampered.xlsx")
            shutil.copyfile(ref_path, tampered_path)

            wb = openpyxl.load_workbook(tampered_path)
            ws = wb.active
            # Ubah satu sel
            orig_val = ws["B2"].value
            ws["B2"].value = str(orig_val) + "_TAMPERED"
            wb.save(tampered_path)

            cmp_res = compare_excel_workbooks(ref_path, tampered_path)
            assert not cmp_res.is_identical
            assert len(cmp_res.differences) > 0
            assert any("B2" in d["location"] or "R2C2" in d["location"] for d in cmp_res.differences)
