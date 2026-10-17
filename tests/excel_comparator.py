import openpyxl
from typing import List, Dict, Any, Tuple


class ExcelComparisonResult:
    def __init__(self):
        self.total_cells: int = 0
        self.differences: List[Dict[str, Any]] = []

    @property
    def is_identical(self) -> bool:
        return len(self.differences) == 0

    def add_diff(self, category: str, location: str, expected: Any, actual: Any, detail: str = ""):
        self.differences.append({
            "category": category,
            "location": location,
            "expected": expected,
            "actual": actual,
            "detail": detail
        })

    def summary(self) -> str:
        res = [f"Total sel diperiksa: {self.total_cells}", f"Jumlah perbedaan: {len(self.differences)}"]
        if self.differences:
            res.append("Detail perbedaan (maksimal 20 pertama):")
            for diff in self.differences[:20]:
                res.append(f"  - [{diff['category']}] Lokasi {diff['location']}: Harapan='{diff['expected']}', Aktual='{diff['actual']}' {diff['detail']}")
        return "\n".join(res)


def compare_excel_workbooks(
    file_expected: str,
    file_actual: str,
    check_styles: bool = True
) -> ExcelComparisonResult:
    """Membandingkan dua file Excel secara logis menggunakan openpyxl.
    
    Pengecualian yang sah:
    - Metadata dokumen/workbook (waktu dibuat, dimodifikasi, pembuat) yang pasti berbeda tiap run.
    """
    result = ExcelComparisonResult()

    wb_exp = openpyxl.load_workbook(file_expected, data_only=False)
    wb_act = openpyxl.load_workbook(file_actual, data_only=False)

    # 1. Sheet names & order
    if wb_exp.sheetnames != wb_act.sheetnames:
        result.add_diff(
            category="SHEET_NAMES",
            location="Workbook",
            expected=wb_exp.sheetnames,
            actual=wb_act.sheetnames,
            detail="Nama atau urutan sheet berbeda."
        )
        return result

    for sheet_name in wb_exp.sheetnames:
        ws_exp = wb_exp[sheet_name]
        ws_act = wb_act[sheet_name]

        # 2. Dimensions
        if ws_exp.max_row != ws_act.max_row or ws_exp.max_column != ws_act.max_column:
            result.add_diff(
                category="DIMENSIONS",
                location=f"Sheet '{sheet_name}'",
                expected=(ws_exp.max_row, ws_exp.max_column),
                actual=(ws_act.max_row, ws_act.max_column),
                detail=f"Dimensi baris/kolom berbeda."
            )

        max_r = max(ws_exp.max_row or 0, ws_act.max_row or 0)
        max_c = max(ws_exp.max_column or 0, ws_act.max_column or 0)

        # 3. Cell values, data types & formatting
        for r in range(1, max_r + 1):
            for c in range(1, max_c + 1):
                result.total_cells += 1
                cell_loc = f"{sheet_name}!R{r}C{c}"
                c_exp = ws_exp.cell(r, c)
                c_act = ws_act.cell(r, c)

                # Value comparison
                v_exp = c_exp.value
                v_act = c_act.value

                if v_exp != v_act:
                    result.add_diff(
                        category="CELL_VALUE",
                        location=cell_loc,
                        expected=v_exp,
                        actual=v_act,
                        detail="Nilai sel tidak sama."
                    )

                # Data type comparison
                if c_exp.data_type != c_act.data_type:
                    result.add_diff(
                        category="DATA_TYPE",
                        location=cell_loc,
                        expected=c_exp.data_type,
                        actual=c_act.data_type,
                        detail="Tipe data sel tidak sama."
                    )

                # Style & format comparison
                if check_styles:
                    if c_exp.number_format != c_act.number_format:
                        result.add_diff(
                            category="NUMBER_FORMAT",
                            location=cell_loc,
                            expected=c_exp.number_format,
                            actual=c_act.number_format,
                            detail="Format angka sel tidak sama."
                        )

        # 4. Freeze panes, Merged cells, Autofilter
        if ws_exp.freeze_panes != ws_act.freeze_panes:
            result.add_diff(
                category="FREEZE_PANES",
                location=f"Sheet '{sheet_name}'",
                expected=ws_exp.freeze_panes,
                actual=ws_act.freeze_panes
            )

        merged_exp = sorted([str(m) for m in ws_exp.merged_cells.ranges])
        merged_act = sorted([str(m) for m in ws_act.merged_cells.ranges])
        if merged_exp != merged_act:
            result.add_diff(
                category="MERGED_CELLS",
                location=f"Sheet '{sheet_name}'",
                expected=merged_exp,
                actual=merged_act
            )

    return result
