"""Data operations: column management, row insertion/deletion, value replacements using pywin32 COM."""

from typing import List, Dict, Any, Optional
from copilots_app.services.excel.com_utils import col_index_to_letter


class DataOps:
    """Executes data level modifications on active Excel Workbook COM object."""

    @classmethod
    def _get_sheet(cls, wb: Any, sheet_name: Optional[str]) -> Any:
        if not sheet_name:
            return wb.ActiveSheet
        for i in range(1, wb.Worksheets.Count + 1):
            if wb.Worksheets(i).Name.lower() == sheet_name.lower():
                return wb.Worksheets(i)
        return wb.ActiveSheet

    @classmethod
    def execute_set_cell_value(
        cls,
        wb: Any,
        sheet_name: str,
        cell_ref: str,
        value: Any,
    ) -> str:
        """Set the value of a specific cell (e.g. B2)."""
        ws = cls._get_sheet(wb, sheet_name)
        ws.Range(cell_ref).Value = value
        return f"Set cell '{cell_ref}' to '{value}' on sheet '{ws.Name}'."

    @classmethod
    def execute_add_column(
        cls,
        wb: Any,
        sheet_name: str,
        column_name: str,
        table_name: Optional[str] = None,
        default_values: Optional[List[Any]] = None,
    ) -> str:
        """Add a new column to a table or worksheet."""
        ws = cls._get_sheet(wb, sheet_name)

        # 1. Native Excel Table (ListObject)
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
                new_col = lo.ListColumns.Add()
                new_col.Name = column_name

                if default_values and hasattr(new_col, "DataBodyRange") and new_col.DataBodyRange:
                    rng = new_col.DataBodyRange
                    num_rows = min(len(default_values), rng.Rows.Count)
                    for idx in range(num_rows):
                        rng.Cells(idx + 1, 1).Value = default_values[idx]

                return f"Added column '{column_name}' to table '{table_name}'."

        # 2. Standard Worksheet Grid
        used_r = ws.UsedRange
        max_col = used_r.Column + used_r.Columns.Count - 1
        new_col_idx = max_col + 1

        # Header in Row 1
        ws.Cells(1, new_col_idx).Value = column_name

        if default_values:
            for idx, val in enumerate(default_values):
                ws.Cells(2 + idx, new_col_idx).Value = val

        return f"Added column '{column_name}' to sheet '{ws.Name}' at column {col_index_to_letter(new_col_idx)}."

    @classmethod
    def execute_remove_column(
        cls,
        wb: Any,
        sheet_name: str,
        column_name: str,
        table_name: Optional[str] = None,
    ) -> str:
        """Remove a column by header name."""
        ws = cls._get_sheet(wb, sheet_name)

        # 1. Check Table (ListObject)
        if table_name:
            try:
                for i in range(1, ws.ListObjects.Count + 1):
                    lo = ws.ListObjects(i)
                    if lo.Name.lower() == table_name.lower():
                        for c in range(1, lo.ListColumns.Count + 1):
                            if lo.ListColumns(c).Name.lower() == column_name.lower():
                                lo.ListColumns(c).Delete()
                                return f"Removed column '{column_name}' from table '{table_name}'."
            except Exception:
                pass

        # 2. Standard Worksheet Grid
        used_r = ws.UsedRange
        start_col = used_r.Column
        cols_count = used_r.Columns.Count

        target_col = None
        for c in range(start_col, start_col + cols_count):
            val = ws.Cells(1, c).Value
            if val and str(val).strip().lower() == column_name.strip().lower():
                target_col = c
                break

        if not target_col:
            raise ValueError(f"Column '{column_name}' not found in sheet '{ws.Name}'.")

        ws.Columns(target_col).Delete()
        return f"Removed column '{column_name}' from '{ws.Name}'."

    @classmethod
    def execute_rename_column(
        cls,
        wb: Any,
        sheet_name: str,
        old_column_name: str,
        new_column_name: str,
    ) -> str:
        """Rename an existing column header."""
        ws = cls._get_sheet(wb, sheet_name)

        # Check tables first
        try:
            for i in range(1, ws.ListObjects.Count + 1):
                lo = ws.ListObjects(i)
                for c in range(1, lo.ListColumns.Count + 1):
                    if lo.ListColumns(c).Name.lower() == old_column_name.lower():
                        lo.ListColumns(c).Name = new_column_name
                        return f"Renamed table column '{old_column_name}' to '{new_column_name}' in table '{lo.Name}'."
        except Exception:
            pass

        # Check row 1 across used range
        used_r = ws.UsedRange
        start_col = used_r.Column
        cols_count = used_r.Columns.Count

        for c in range(start_col, start_col + cols_count):
            cell = ws.Cells(1, c)
            if cell.Value and str(cell.Value).strip().lower() == old_column_name.strip().lower():
                cell.Value = new_column_name
                return f"Renamed column '{old_column_name}' to '{new_column_name}' in sheet '{ws.Name}'."

        raise ValueError(f"Column '{old_column_name}' not found in '{ws.Name}'.")

    @classmethod
    def execute_replace_values(
        cls,
        wb: Any,
        sheet_name: str,
        old_val: Any,
        new_val: Any,
    ) -> str:
        """Find and replace all instances of matching value in sheet using COM Replace."""
        ws = cls._get_sheet(wb, sheet_name)

        try:
            # Excel Range.Replace: (What, Replacement, LookAt, SearchOrder, MatchCase)
            # xlPart = 2, xlWhole = 1
            ws.UsedRange.Replace(
                What=str(old_val),
                Replacement=str(new_val),
                LookAt=1,  # xlWhole
                SearchOrder=1,  # xlByRows
                MatchCase=False,
            )
            return f"Replaced values matching '{old_val}' with '{new_val}' in '{ws.Name}'."
        except Exception as e:
            raise ValueError(f"Failed to replace values in '{ws.Name}': {e}")
