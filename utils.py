"""
Finextract Utilities Module
Fungsi-fungsi pembantu umum untuk integrasi sistem operasi,
pembersihan teks, dan transformasi DataFrame.
"""
import os
import sys
import re
import subprocess
import logging
from typing import Any
import pandas as pd

logger = logging.getLogger(__name__)


def open_system_path(path: str) -> bool:
    """Membuka file atau folder menggunakan aplikasi default OS (Windows, macOS, Linux)."""
    if not path or not os.path.exists(path):
        return False
    try:
        if sys.platform == "win32":
            os.startfile(path)
        elif sys.platform == "darwin":
            subprocess.Popen(["open", path])
        else:
            subprocess.Popen(["xdg-open", path])
        return True
    except Exception as e:
        logger.warning("Gagal membuka path sistem '%s': %s", path, e)
        return False


def clean_text(text: Any) -> str:
    """Membersihkan teks dari newline dan spasi berlebih."""
    if text is None:
        return ""
    text_str = str(text).replace("\n", " ")
    return re.sub(r"\s+", " ", text_str).strip()


def swap_columns(df: pd.DataFrame, col_a: str, col_b: str) -> pd.DataFrame:
    """Menukar nilai antara dua kolom dalam DataFrame secara in-place dan aman."""
    if col_a in df.columns and col_b in df.columns:
        temp = df[col_a].copy()
        df[col_a] = df[col_b]
        df[col_b] = temp
    return df
