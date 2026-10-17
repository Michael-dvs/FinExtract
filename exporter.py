import os
import pandas as pd
from typing import Optional, List
from settings import Settings, SettingsManager
from formatters import format_number_value, format_date_value


def identify_amount_columns(columns: List[str]) -> List[str]:
    """Mengidentifikasi kolom mutasi angka berdasarkan nama kolom."""
    amount_keywords = ("debit", "credit", "balance", "saldo", "debet", "kredit", "amount")
    return [col for col in columns if any(kw in col.lower() for kw in amount_keywords)]


def identify_date_columns(columns: List[str]) -> List[str]:
    """Mengidentifikasi kolom tanggal berdasarkan nama kolom."""
    date_keywords = ("date", "tanggal", "tgl")
    return [col for col in columns if any(kw in col.lower() for kw in date_keywords)]


def export_dataframe_to_excel(
    df: pd.DataFrame,
    output_excel_path: str,
    settings: Optional[Settings] = None
) -> int:
    """Mengekspor DataFrame mutasi bank ke berkas Excel dengan pemformatan penyajian terpadu."""
    if df is None or len(df) == 0:
        return 0

    if settings is None:
        settings = SettingsManager().load()

    df_out = df.copy()

    amt_cols = identify_amount_columns(list(df_out.columns))
    date_cols = identify_date_columns(list(df_out.columns))

    # 1. Pemformatan Tanggal (hanya jika pengguna memilih non-default)
    if settings.date_format != "DD/MM/YYYY":
        for col in date_cols:
            df_out[col] = df_out[col].apply(lambda v: format_date_value(v, settings.date_format))

    # 2. Pemformatan Angka
    # Jika mode numerik:
    if settings.output_number_type == "numeric":
        for col in amt_cols:
            df_out[col] = df_out[col].apply(lambda v: format_number_value(v, settings))

        try:
            with pd.ExcelWriter(output_excel_path, engine="openpyxl") as writer:
                df_out.to_excel(writer, index=False)
                ws = writer.sheets["Sheet1"] if "Sheet1" in writer.sheets else writer.book.active

                # Format invarian OpenXML (#,##0.00 atau #,##0)
                num_fmt = "#,##0" if settings.decimal_places == 0 else f"#,##0.{'0' * settings.decimal_places}"

                col_indices = [df_out.columns.get_loc(c) + 1 for c in amt_cols]
                for row in ws.iter_rows(min_row=2, max_row=ws.max_row):
                    for col_idx in col_indices:
                        cell = row[col_idx - 1]
                        if isinstance(cell.value, (int, float)):
                            cell.number_format = num_fmt
            return len(df_out)
        except PermissionError:
            print(f"ERROR: Gagal menyimpan file ke '{output_excel_path}'.")
            print("SOLUSI: Pastikan file Excel tersebut TIDAK SEDANG DIBUKA. Tutup file Excel lalu coba lagi.")
            return 0
        except Exception as e:
            print(f"Gagal menyimpan file Excel numerik: {e}")
            return 0

    # Mode Teks (default):
    is_default_num = (
        settings.thousands_separator == ","
        and settings.decimal_separator == "."
        and settings.decimal_places == 2
        and settings.negative_format == "-1234"
    )
    if not is_default_num:
        for col in amt_cols:
            df_out[col] = df_out[col].apply(lambda v: format_number_value(v, settings))

    try:
        df_out.to_excel(output_excel_path, index=False)
        return len(df_out)
    except PermissionError:
        print(f"ERROR: Gagal menyimpan file ke '{output_excel_path}'.")
        print("SOLUSI: Pastikan file Excel tersebut TIDAK SEDANG DIBUKA. Tutup file Excel lalu coba lagi.")
        return 0
    except Exception as e:
        print(f"Gagal menyimpan file Excel: {e}")
        return 0
