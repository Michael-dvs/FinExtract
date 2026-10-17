import re
import os
from typing import List, Tuple, Callable, Optional, Any
from PIL import Image
import numpy as np

# Lazy imports for heavy packages
_ocr_engine_instance = None


def get_ocr_engine():
    """Mengembalikan instance singleton RapidOCR."""
    global _ocr_engine_instance
    if _ocr_engine_instance is None:
        try:
            from rapidocr_onnxruntime import RapidOCR
            _ocr_engine_instance = RapidOCR()
        except ImportError as e:
            raise RuntimeError(
                f"Engine OCR 'rapidocr-onnxruntime' belum terpasang: {e}. "
                "Silahkan jalankan: pip install rapidocr-onnxruntime"
            )
    return _ocr_engine_instance


def is_ocr_available() -> bool:
    """Memeriksa ketersediaan library OCR dan renderer PDF."""
    try:
        import pypdfium2
        import onnxruntime
        from rapidocr_onnxruntime import RapidOCR
        return True
    except ImportError:
        return False


def normalize_ocr_text(text: str) -> str:
    """Menormalkan teks hasil OCR.
    
    Menyisipkan spasi pada tanggal yang menempel rapat (misal: '03Jun2026' -> '03 Jun 2026').
    """
    if not text:
        return ""
    s = str(text).strip()
    s = re.sub(r"^(\d{1,2})([A-Za-z]{3})", r"\1 \2", s)
    s = re.sub(r"^(\d{1,2}\s+[A-Za-z]{3})(\d{4})", r"\1 \2", s)
    return s


# Profil fraksi kolom standar untuk Mandiri / Livin (relatif 0.0 - 1.0 terhadap lebar halaman)
LIVIN_COLUMNS_FRACTION = [
    (0.000, 0.185),  # 0: Date
    (0.185, 0.385),  # 1: Remark
    (0.385, 0.570),  # 2: Ref No
    (0.570, 0.715),  # 3: Debit
    (0.715, 0.860),  # 4: Credit
    (0.860, 1.000)   # 5: Balance
]


def extract_table_grid_from_image(
    pil_img: Image.Image,
    columns_fraction: List[Tuple[float, float]],
    header_keywords: Optional[List[str]] = None,
    engine: Any = None
) -> List[List[str]]:
    """Mengekstrak baris tabel mentah (raw_rows) dari gambar halaman rekening koran.
    
    Mengelompokkan teks berdasarkan overlap sumbu Y dan memproyeksikannya ke kolom sumbu X.
    """
    if engine is None:
        engine = get_ocr_engine()

    img_rgb = pil_img.convert("RGB")
    w, h = img_rgb.size
    results, _ = engine(np.array(img_rgb))
    if not results:
        return []

    if header_keywords is None:
        header_keywords = ["posting date", "keterangan", "tanggal"]

    # Deteksi batas vertikal tabel (header dan footer)
    header_y = 0.0
    for box, text, _ in results:
        t_lower = text.lower()
        if any(kw in t_lower for kw in header_keywords):
            header_y = max(header_y, box[2][1])

    # Jika header berulang tidak ditemukan pada halaman lanjutan, gunakan margin atas wajar
    if header_y == 0.0:
        header_y = h * 0.06

    footer_y = h * 0.94

    # Saring kotak teks dalam area tabel
    data_boxes = []
    for box, text, score in results:
        y_center = (box[0][1] + box[2][1]) / 2.0
        x_center = (box[0][0] + box[1][0]) / 2.0
        if header_y < y_center < footer_y:
            data_boxes.append((y_center, x_center, text.strip(), box))

    if not data_boxes:
        return []

    data_boxes.sort(key=lambda b: b[0])

    # Pengelompokan baris deterministik berdasarkan overlap vertikal
    lines = []
    for b in data_boxes:
        placed = False
        b_h = max(1.0, b[3][2][1] - b[3][0][1])
        for line in lines:
            avg_y = sum(item[0] for item in line) / len(line)
            if abs(b[0] - avg_y) <= (b_h * 0.5):
                line.append(b)
                placed = True
                break
        if not placed:
            lines.append([b])

    lines.sort(key=lambda l: sum(item[0] for item in l) / len(l))

    # Petakan setiap baris ke kolom
    num_cols = len(columns_fraction)
    raw_rows = []

    for line in lines:
        line.sort(key=lambda b: b[1])
        row = [""] * num_cols
        for b in line:
            x_rel = b[1] / w
            col_idx = 0
            for i, (xmin, xmax) in enumerate(columns_fraction):
                if xmin <= x_rel < xmax:
                    col_idx = i
                    break
            text_val = b[2]
            if col_idx == 0:
                text_val = normalize_ocr_text(text_val)

            if row[col_idx]:
                row[col_idx] += " " + text_val
            else:
                row[col_idx] = text_val

        if any(row):
            raw_rows.append(row)

    return raw_rows


def extract_pdf_ocr_raw_rows(
    pdf_path: str,
    columns_fraction: Optional[List[Tuple[float, float]]] = None,
    dpi: int = 200,
    progress_callback: Optional[Callable[[int, int, str], None]] = None
) -> List[List[str]]:
    """Merender halaman PDF menjadi gambar dan mengekstrak tabel raw_rows dari seluruh halaman."""
    try:
        import pypdfium2 as pdfium
    except ImportError as e:
        raise RuntimeError(f"Library 'pypdfium2' diperlukan untuk merender PDF: {e}")

    if columns_fraction is None:
        columns_fraction = LIVIN_COLUMNS_FRACTION

    engine = get_ocr_engine()
    doc = pdfium.PdfDocument(pdf_path)
    total_pages = len(doc)
    scale = dpi / 72.0

    all_raw_rows = []

    for page_idx in range(total_pages):
        page_num = page_idx + 1
        if progress_callback:
            progress_callback(page_num, total_pages, f"Memproses OCR halaman {page_num}/{total_pages}...")

        page = doc[page_idx]
        bitmap = page.render(scale=scale)
        pil_img = bitmap.to_pil().convert("RGB")

        page_rows = extract_table_grid_from_image(pil_img, columns_fraction, engine=engine)
        all_raw_rows.extend(page_rows)

    return all_raw_rows


def validate_running_balance(
    df: Any,
    col_debit: str = "Debit",
    col_credit: str = "Credit",
    col_balance: str = "Balance"
) -> Tuple[bool, List[str]]:
    """Memeriksa konsistensi matematis saldo berjalan pada DataFrame mutasi.
    
    Saldo_n = Saldo_{n-1} + Mutasi
    Mengembalikan (is_valid, daftar_peringatan).
    """
    if df is None or len(df) <= 1:
        return True, []

    for c in [col_debit, col_credit, col_balance]:
        if c not in df.columns:
            return True, [f"Kolom '{c}' tidak ditemukan untuk validasi saldo."]

    def parse_amount(val: Any) -> float:
        if val is None or val == "" or str(val).strip() == "-":
            return 0.0
        s = str(val).replace(",", "").strip()
        try:
            return float(s)
        except ValueError:
            return 0.0

    warnings = []
    inconsistent_count = 0

    for i in range(1, len(df)):
        bal_prev = parse_amount(df.iloc[i - 1][col_balance])
        bal_curr = parse_amount(df.iloc[i][col_balance])
        amt_d = parse_amount(df.iloc[i][col_debit])
        amt_c = parse_amount(df.iloc[i][col_credit])

        # FinExtract menukar posisi Debit dan Credit di tahap akhir exporter.
        # Kita uji kedua kemungkinan konsistensi:
        diff1 = round(bal_prev + amt_c - amt_d - bal_curr, 2)
        diff2 = round(bal_prev - amt_c + amt_d - bal_curr, 2)

        if diff1 != 0.0 and diff2 != 0.0:
            inconsistent_count += 1
            if len(warnings) < 10:
                warnings.append(
                    f"Baris {i+1}: Selisih saldo terdeteksi (Saldo Sebelumnya: {bal_prev:,.2f}, "
                    f"Mutasi 1: {amt_d:,.2f}, Mutasi 2: {amt_c:,.2f}, Saldo Tercatat: {bal_curr:,.2f})."
                )

    if inconsistent_count > 0:
        return False, warnings
    return True, []
