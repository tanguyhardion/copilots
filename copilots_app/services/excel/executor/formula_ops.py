"""Formula operations: insert, replace, fill semantic formulas across worksheet columns using pywin32 COM."""

from typing import Optional, Dict, Any
from copilots_app.services.excel.com_utils import col_index_to_letter
from copilots_app.services.excel.utils.formula_translator import FormulaTranslator


class FormulaOps:
    """Executes formula insertions and translations on active Excel Workbook COM object."""

    @classmethod
    def _get_sheet(cls, wb: Any, sheet_name: Optional[str]) -> Any:
        if not sheet_name:
            return wb.ActiveSheet
        for i in range(1, wb.Worksheets.Count + 1):
            if wb.Worksheets(i).Name.lower() == sheet_name.lower():
                return wb.Worksheets(i)
        return wb.ActiveSheet

    @classmethod
    def execute_insert_formula(
        cls,
        wb: Any,
        sheet_name: str,
        column_name: str,
        semantic_formula: str,
        table_name: Optional[str] = None,
    ) -> str:
        """Insert semantic formula down a target column."""
        ws = cls._get_sheet(wb, sheet_name)

        # 1. Native Table (ListObject)
        if table_name:
            lo = None
            try:
                for i in range(1, ws.ListObjects.Count + 1):
                    if ws.ListObjects(i).Name.lower() == table_name.lower():
                        lo = ws.ListObjects(i)
                        break
            except Exception:
                pass

            if lo:
                # Find or add column
                target_col = None
                for c in range(1, lo.ListColumns.Count + 1):
                    if lo.ListColumns(c).Name.lower() == column_name.lower():
                        target_col = lo.ListColumns(c)
                        break

                if not target_col:
                    target_col = lo.ListColumns.Add()
                    target_col.Name = column_name

                # Build column map
                col_map: Dict[str, str] = {}
                for c in range(1, lo.ListColumns.Count + 1):
                    c_name = lo.ListColumns(c).Name
                    col_letter = col_index_to_letter(lo.Range.Column + c - 1)
                    col_map[c_name.lower()] = col_letter

                excel_formula = FormulaTranslator.translate(
                    semantic_formula=semantic_formula,
                    col_name_to_letter=col_map,
                    is_native_table=True,
                )

                if hasattr(target_col, "DataBodyRange") and target_col.DataBodyRange:
                    target_col.DataBodyRange.Formula = excel_formula
                    rows_count = target_col.DataBodyRange.Rows.Count
                else:
                    rows_count = 0

                return f"Inserted formula '{excel_formula}' into table '{table_name}' column '{column_name}' ({rows_count} rows)."

        # 2. Standard Worksheet Grid
        used_r = ws.UsedRange
        min_r = used_r.Row
        max_r = used_r.Row + used_r.Rows.Count - 1
        min_c = used_r.Column
        max_c = used_r.Column + used_r.Columns.Count - 1

        col_map = {}
        target_col_idx = None

        for c in range(min_c, max_c + 1):
            val = ws.Cells(1, c).Value
            if val is not None:
                header_name = str(val).strip()
                col_letter = col_index_to_letter(c)
                col_map[header_name.lower()] = col_letter
                if header_name.lower() == column_name.strip().lower():
                    target_col_idx = c

        if not target_col_idx:
            target_col_idx = max_c + 1
            ws.Cells(1, target_col_idx).Value = column_name
            col_map[column_name.strip().lower()] = col_index_to_letter(target_col_idx)

        start_row = 2
        end_row = max(start_row, max_r)
        inserted_count = 0

        for r in range(start_row, end_row + 1):
            excel_formula = FormulaTranslator.translate(
                semantic_formula=semantic_formula,
                col_name_to_letter=col_map,
                row_index=r,
                is_native_table=False,
            )
            ws.Cells(r, target_col_idx).Formula = excel_formula
            inserted_count += 1

        return f"Inserted formula '{semantic_formula}' into column '{column_name}' across {inserted_count} row(s)."
