"""Sheet operations: create, rename, duplicate, delete worksheets using Excel COM (pywin32)."""

from typing import Any


class SheetOps:
    """Executes worksheet level operations on active Excel Workbook COM object."""

    @classmethod
    def execute_create_sheet(cls, wb: Any, sheet_name: str) -> str:
        """Create a new worksheet in the active workbook."""
        existing_names = [wb.Worksheets(i).Name.lower() for i in range(1, wb.Worksheets.Count + 1)]

        target_name = sheet_name
        if target_name.lower() in existing_names:
            counter = 1
            new_name = f"{sheet_name}_{counter}"
            while new_name.lower() in existing_names:
                counter += 1
                new_name = f"{sheet_name}_{counter}"
            target_name = new_name

        # Add after last sheet
        last_sheet = wb.Worksheets(wb.Worksheets.Count)
        new_ws = wb.Worksheets.Add(After=last_sheet)
        new_ws.Name = target_name
        return f"Created worksheet '{target_name}'."

    @classmethod
    def execute_rename_sheet(cls, wb: Any, old_name: str, new_name: str) -> str:
        """Rename an existing worksheet."""
        ws = None
        for i in range(1, wb.Worksheets.Count + 1):
            if wb.Worksheets(i).Name.lower() == old_name.lower():
                ws = wb.Worksheets(i)
                break

        if not ws:
            raise ValueError(f"Worksheet '{old_name}' not found in workbook.")

        ws.Name = new_name
        return f"Renamed worksheet '{old_name}' to '{new_name}'."

    @classmethod
    def execute_duplicate_sheet(cls, wb: Any, sheet_name: str, new_name: str = None) -> str:
        """Duplicate an existing worksheet."""
        ws = None
        for i in range(1, wb.Worksheets.Count + 1):
            if wb.Worksheets(i).Name.lower() == sheet_name.lower():
                ws = wb.Worksheets(i)
                break

        if not ws:
            raise ValueError(f"Worksheet '{sheet_name}' not found.")

        ws.Copy(After=ws)
        # The copy is now active sheet
        dup_ws = wb.ActiveSheet
        if new_name:
            dup_ws.Name = new_name

        return f"Duplicated worksheet '{sheet_name}' as '{dup_ws.Name}'."

    @classmethod
    def execute_delete_sheet(cls, wb: Any, sheet_name: str) -> str:
        """Delete a worksheet from workbook."""
        if wb.Worksheets.Count <= 1:
            raise ValueError("Cannot delete the only remaining sheet in the workbook.")

        ws = None
        for i in range(1, wb.Worksheets.Count + 1):
            if wb.Worksheets(i).Name.lower() == sheet_name.lower():
                ws = wb.Worksheets(i)
                break

        if not ws:
            raise ValueError(f"Worksheet '{sheet_name}' not found.")

        # Suppress Excel warning prompt when deleting sheet
        app = wb.Application
        orig_alerts = app.DisplayAlerts
        try:
            app.DisplayAlerts = False
            ws.Delete()
        finally:
            app.DisplayAlerts = orig_alerts

        return f"Deleted worksheet '{sheet_name}'."
