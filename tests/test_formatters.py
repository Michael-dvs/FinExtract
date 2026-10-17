import pytest
from settings import Settings
from formatters import (
    parse_raw_amount,
    format_number_value,
    format_date_value,
    format_output_filename,
    mask_sensitive_text,
)


class TestFormatters:
    """Rangkaian pengujian table-driven untuk modul formatters."""

    @pytest.mark.parametrize(
        "raw_input, expected",
        [
            ("1,234,567.89", 1234567.89),
            ("1.234.567,89", 1234567.89),
            ("368,500.00", 368500.00),
            ("0.00", 0.0),
            (" 0.00", 0.0),
            ("-500.50", -500.50),
            ("(1,000.00)", -1000.00),
            ("1000-", -1000.00),
            ("-", None),
            ("", None),
            (None, None),
            (12345.67, 12345.67),
        ],
    )
    def test_parse_raw_amount(self, raw_input, expected):
        res = parse_raw_amount(raw_input)
        if expected is None:
            assert res is None
        else:
            assert pytest.approx(res, rel=1e-5) == expected

    def test_format_number_default_settings(self):
        """Pada pengaturan default, keluaran harus identik dengan format baseline (ribuan ',', desimal '.')."""
        default_settings = Settings()
        assert format_number_value("1234567.89", default_settings) == "1,234,567.89"
        assert format_number_value("0.00", default_settings) == "0.00"
        assert format_number_value("-500.00", default_settings) == "-500.00"
        assert format_number_value("-", default_settings) == "-"

    def test_format_number_indonesian_settings(self):
        """Format gaya Indonesia: ribuan '.', desimal ',', negatif dalam kurung."""
        indo_settings = Settings(
            thousands_separator=".",
            decimal_separator=",",
            decimal_places=2,
            negative_format="(1234)",
            output_number_type="text",
        )
        assert format_number_value("1234567.89", indo_settings) == "1.234.567,89"
        assert format_number_value("0.00", indo_settings) == "0,00"
        assert format_number_value("-500.00", indo_settings) == "(500,00)"

    def test_format_number_numeric_mode(self):
        """Mode numerik harus mengembalikan tipe float murni."""
        num_settings = Settings(output_number_type="numeric")
        res = format_number_value("1,234,567.89", num_settings)
        assert isinstance(res, float)
        assert res == 1234567.89

        res_zero = format_number_value("0.00", num_settings)
        assert isinstance(res_zero, float)
        assert res_zero == 0.0

    @pytest.mark.parametrize(
        "date_input, target_fmt, expected",
        [
            ("01/11/2025 14:23:45", "DD/MM/YYYY", "01/11/2025"),
            ("01/11/2025 14:23:45", "YYYY-MM-DD", "2025-11-01"),
            ("01 Dec 2025, 10:00", "DD/MM/YYYY", "01/12/2025"),
            ("01 Dec 2025, 10:00", "YYYY-MM-DD", "2025-12-01"),
            ("03/06/26", "DD/MM/YYYY", "03/06/2026"),
            ("2026-06-03", "DD/MM/YYYY", "03/06/2026"),
            ("2026-06-03", "DD MMM YYYY", "03 Jun 2026"),
            ("2026-06-03", "DD-MM-YYYY", "03-06-2026"),
        ],
    )
    def test_format_date_value(self, date_input, target_fmt, expected):
        assert format_date_value(date_input, target_fmt) == expected

    def test_format_output_filename(self):
        pdf = "/path/to/Mandiri_Statement.pdf"
        assert format_output_filename("{nama_pdf}", pdf) == "Mandiri_Statement.xlsx"
        assert format_output_filename("{nama_pdf}_{bank}", pdf, "Mandiri") == "Mandiri_Statement_Mandiri.xlsx"
        res_date = format_output_filename("{nama_pdf}_{tanggal}", pdf)
        assert res_date.startswith("Mandiri_Statement_")
        assert res_date.endswith(".xlsx")

    def test_mask_sensitive_text(self):
        text = "Rekening 20390100078356499102 berhasil diproses"
        masked = mask_sensitive_text(text)
        assert "20390100078356499102" not in masked
        assert "2039" in masked
        assert "9102" in masked
        assert "******" in masked
