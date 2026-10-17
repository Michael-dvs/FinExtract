import os
import re
from datetime import datetime
from typing import Any, Optional
from settings import Settings


def parse_raw_amount(val: Any) -> Optional[float]:
    """Mengubah nilai mentah (string/number) dari rekening koran menjadi float invarian.
    
    Menangani berbagai format bank: '1,234,567.89', '1.234.567,89', '-500.00', ' 0.00', '-', dsb.
    """
    if val is None:
        return None
    if isinstance(val, (int, float)):
        return float(val)

    s = str(val).strip()
    if not s or s == "-" or s == "":
        return None

    # Hapus spasi dan tanda kurung negatif jika ada (misal '(1,234.50)')
    is_negative = False
    if s.startswith("(") and s.endswith(")"):
        is_negative = True
        s = s[1:-1].strip()
    elif s.startswith("-"):
        is_negative = True
        s = s[1:].strip()
    elif s.endswith("-"):
        is_negative = True
        s = s[:-1].strip()

    # Deteksi pola desimal vs ribuan:
    # 1. Jika ada koma dan titik:
    if "," in s and "." in s:
        last_comma = s.rfind(",")
        last_dot = s.rfind(".")
        if last_comma > last_dot:
            # Format Eropa/Indonesia: 1.234.567,89
            s = s.replace(".", "").replace(",", ".")
        else:
            # Format US/Standar: 1,234,567.89
            s = s.replace(",", "")
    elif "," in s:
        # Hanya ada koma: misal '1234,50' atau '1,234'
        parts = s.split(",")
        if len(parts) == 2 and len(parts[1]) in (1, 2):
            # Desimal koma: '1234,50'
            s = s.replace(",", ".")
        else:
            # Ribuan koma tanpa desimal: '1,234'
            s = s.replace(",", "")
    elif "." in s:
        # Hanya ada titik: misal '1234.50' atau '1.234'
        parts = s.split(".")
        if len(parts) == 2 and len(parts[1]) in (1, 2):
            # Desimal titik: '1234.50'
            pass
        else:
            # Ribuan titik: '1.234.567'
            s = s.replace(".", "")

    try:
        num = float(s)
        return -num if is_negative else num
    except ValueError:
        return None


def format_number_value(val: Any, settings: Settings) -> Any:
    """Memformat nilai angka sesuai pengaturan (Dual-Mode: Teks terformat vs Numerik invarian)."""
    parsed = parse_raw_amount(val)
    if parsed is None:
        # Jika bukan angka (misal tanda '-' atau teks kosong), kembalikan apa adanya
        return val if val is not None else ""

    # Mode Numerik: Kembalikan float murni untuk perhitungan formula Excel
    if settings.output_number_type == "numeric":
        return round(parsed, settings.decimal_places)

    # Mode Teks: Bangun representasi string terformat
    is_neg = parsed < 0
    abs_val = abs(parsed)

    # Format angka dengan presisi desimal
    fmt_spec = f"{{:.{settings.decimal_places}f}}"
    base_str = fmt_spec.format(abs_val)

    int_part, dec_part = base_str.split(".") if "." in base_str else (base_str, "")

    # Tambahkan pemisah ribuan
    if settings.thousands_separator:
        # Sisipkan pemisah ribuan dari kanan ke kiri
        rev_int = int_part[::-1]
        chunks = [rev_int[i:i + 3] for i in range(0, len(rev_int), 3)]
        int_formatted = settings.thousands_separator.join(chunks)[::-1]
    else:
        int_formatted = int_part

    # Gabungkan dengan pemisah desimal
    if settings.decimal_places > 0 and dec_part:
        res = f"{int_formatted}{settings.decimal_separator}{dec_part}"
    else:
        res = int_formatted

    # Format tanda negatif
    if is_neg:
        if settings.negative_format == "(1234)":
            res = f"({res})"
        else:
            res = f"-{res}"

    return res


def format_date_value(date_str: Any, target_format: str = "DD/MM/YYYY") -> str:
    """Mengubah format tanggal string ke format target yang diinginkan."""
    if not date_str:
        return ""

    s = str(date_str).replace("\n", " ").strip()
    # Pisahkan jam jika ada (misal '01/11/2025 14:23:45' atau '01 Dec 2025, 10:00')
    date_clean = s.split(",")[0].strip()

    dt_obj = None

    # Coba berbagai pola tanggal rekening koran
    patterns = [
        ("%d/%m/%Y", r"^\d{1,2}/\d{1,2}/\d{4}"),
        ("%d-%m-%Y", r"^\d{1,2}-\d{1,2}-\d{4}"),
        ("%Y-%m-%d", r"^\d{4}-\d{1,2}-\d{1,2}"),
        ("%d %b %Y", r"^\d{1,2}\s+[A-Za-z]{3}\s+\d{4}"),
        ("%d %B %Y", r"^\d{1,2}\s+[A-Za-z]+\s+\d{4}"),
        ("%d/%m/%y", r"^\d{1,2}/\d{1,2}/\d{2}"),
        ("%d-%m-%y", r"^\d{1,2}-\d{1,2}-\d{2}"),
    ]

    for fmt, regex in patterns:
        m = re.search(regex, date_clean)
        if m:
            matched_str = m.group(0)
            try:
                dt_obj = datetime.strptime(matched_str, fmt)
                break
            except ValueError:
                continue

    if not dt_obj:
        return s

    if target_format == "YYYY-MM-DD":
        return dt_obj.strftime("%Y-%m-%d")
    elif target_format == "DD-MM-YYYY":
        return dt_obj.strftime("%d-%m-%Y")
    elif target_format == "DD MMM YYYY":
        return dt_obj.strftime("%d %b %Y")
    else:  # Default DD/MM/YYYY
        return dt_obj.strftime("%d/%m/%Y")


def format_output_filename(pattern: str, pdf_path: str, bank_name: str = "") -> str:
    """Menghasilkan nama file output Excel berdasarkan pola konfigurasi."""
    base_name = os.path.splitext(os.path.basename(pdf_path))[0]
    today_str = datetime.now().strftime("%Y%m%d")

    filename = pattern.replace("{nama_pdf}", base_name)
    filename = filename.replace("{bank}", bank_name.strip() or "BANK")
    filename = filename.replace("{tanggal}", today_str)

    # Bersihkan karakter terlarang nama file
    filename = re.sub(r'[\\/*?:"<>|]', "", filename).strip()
    if not filename.lower().endswith(".xlsx"):
        filename += ".xlsx"

    return filename


def mask_sensitive_text(text: str) -> str:
    """Menyamarkan nomor rekening panjang dan nominal saldo pada pesan log untuk privasi."""
    if not text:
        return ""

    # Samarkan rangkaian 8 digit atau lebih (misal nomor rekening / kartu)
    def repl_digits(m):
        d = m.group(0)
        if len(d) >= 8:
            return d[:4] + ("*" * (len(d) - 8)) + d[-4:]
        return d

    masked = re.sub(r"\b\d{8,}\b", repl_digits, str(text))

    # Samarkan nominal saldo/mata uang seperti Rp 50.000.000,00 atau IDR 1,234.56
    masked = re.sub(
        r"(?i)\b(Rp\.?|IDR|Saldo:?)\s*([0-9.,]+)",
        r"\1 [DISAMARKAN]",
        masked
    )

    return masked
