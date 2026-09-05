"""Formatting & Style operations: autofit columns, styling headers, freeze panes using pywin32 COM."""

from typing import Optional, Any
from copilots_app.services.excel.com_utils import hex_to_bgr_int


class StyleOps:
    """Executes formatting and visual styling on active Excel Workbook COM object."""

    @classmethod
    def _get_sheet(cls, wb: Any, sheet_name: Optional[str]) -> Any:
        if not sheet_name:
            return wb.ActiveSheet
        for i in range(1, wb.Worksheets.Count + 1):
            if wb.Worksheets(i).Name.lower() == sheet_name.lower():
                return wb.Worksheets(i)
        return wb.ActiveSheet

    @classmethod
    def execute_autofit_columns(cls, wb: Any, sheet_name: str) -> str:
        """Autofit all column widths in a worksheet."""
        ws = cls._get_sheet(wb, sheet_name)
        ws.Columns.AutoFit()
        return f"Autofitted columns for sheet '{ws.Name}'."

    @classmethod
    def execute_apply_style(
        cls,
        wb: Any,
        sheet_name: str,
        target: str = "headers",
        range_ref: Optional[str] = None,
        bg_color: Optional[str] = "1F4E78",  # Default dark blue header
        text_color: Optional[str] = "FFFFFF",
        number_format: Optional[str] = None,
        bold: bool = True,
    ) -> str:
        """Apply visual formatting to headers, cell ranges, or data cells."""
        ws = cls._get_sheet(wb, sheet_name)

        if range_ref:
            target_rng = ws.Range(range_ref)
            if bold is not None:
                target_rng.Font.Bold = bold
            if text_color:
                target_rng.Font.Color = hex_to_bgr_int(text_color)
            if bg_color:
                target_rng.Interior.Color = hex_to_bgr_int(bg_color)
            if number_format:
                target_rng.NumberFormat = number_format
            return f"Applied formatting to range '{range_ref}' on sheet '{ws.Name}'."

        if target == "headers":
            used_r = ws.UsedRange
            start_col = used_r.Column
            cols_count = used_r.Columns.Count

            header_rng = ws.Range(ws.Cells(1, start_col), ws.Cells(1, start_col + cols_count - 1))
            header_rng.Font.Bold = bold
            if text_color:
                header_rng.Font.Color = hex_to_bgr_int(text_color)
            if bg_color:
                header_rng.Interior.Color = hex_to_bgr_int(bg_color)

            # xlHAlignCenter = -4108, xlVAlignCenter = -4108
            header_rng.HorizontalAlignment = -4108
            header_rng.VerticalAlignment = -4108

            return f"Applied header styling to sheet '{ws.Name}'."

        elif number_format:
            used_r = ws.UsedRange
            start_col = used_r.Column
            cols_count = used_r.Columns.Count
            rows_count = used_r.Rows.Count
            if rows_count > 1:
                data_rng = ws.Range(ws.Cells(2, start_col), ws.Cells(rows_count, start_col + cols_count - 1))
                data_rng.NumberFormat = number_format

            return f"Applied number format '{number_format}' to '{ws.Name}' data cells."

        return f"Applied custom style to '{ws.Name}'."

    @classmethod
    def execute_freeze_panes(cls, wb: Any, sheet_name: str, cell_ref: str = "A2") -> str:
        """Freeze panes at target cell reference."""
        ws = cls._get_sheet(wb, sheet_name)
        ws.Activate()
        app = wb.Application
        app.ActiveWindow.FreezePanes = False
        target_cell = ws.Range(cell_ref)
        target_cell.Select()
        app.ActiveWindow.FreezePanes = True

        return f"Frozen panes at '{cell_ref}' on sheet '{ws.Name}'."
