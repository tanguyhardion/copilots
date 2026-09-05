"""Workbook Analyzer: Analyzes active Excel workbooks via pywin32 COM to extract semantic models and generate LLM prompts."""

import os
import json
from typing import Optional, List, Dict, Any

from copilots_app.services.excel.com_utils import (
    get_active_excel_and_wb,
    col_index_to_letter,
)
from copilots_app.services.excel.models.semantic import (
    WorkbookModel,
    WorksheetModel,
    TableModel,
    ColumnModel,
    NamedRangeModel,
    ChartModel,
)


class WorkbookAnalyzer:
    """Extracts semantic structure from an active Excel workbook via pywin32 COM."""

    @classmethod
    def analyze(cls, file_path_or_wb: Optional[Any] = None) -> WorkbookModel:
        """Parse an active Excel workbook via COM and return its complete WorkbookModel.

        Args:
            file_path_or_wb: Optional COM Workbook object or file path string. If None,
                             attaches to the currently active Excel workbook.
        """
        app = None
        wb = None

        if hasattr(file_path_or_wb, "Worksheets"):
            wb = file_path_or_wb
            app = wb.Application
        elif isinstance(file_path_or_wb, str) and file_path_or_wb.strip():
            # If path provided, find matching open workbook or attach
            try:
                from copilots_app.services.excel.com_utils import get_active_excel_app
                app = get_active_excel_app()
                norm_target = os.path.normpath(file_path_or_wb).lower()
                for open_wb in app.Workbooks:
                    if os.path.normpath(open_wb.FullName).lower() == norm_target or open_wb.Name.lower() == os.path.basename(file_path_or_wb).lower():
                        wb = open_wb
                        break
                if not wb:
                    wb = app.Workbooks.Open(os.path.abspath(file_path_or_wb))
            except Exception:
                app, wb = get_active_excel_and_wb()
        else:
            app, wb = get_active_excel_and_wb()

        filename = wb.Name
        file_path = wb.FullName if hasattr(wb, "FullName") else filename

        worksheets: List[WorksheetModel] = []
        for index in range(1, wb.Worksheets.Count + 1):
            ws = wb.Worksheets(index)
            sheet_model = cls._analyze_worksheet(ws, index - 1)
            worksheets.append(sheet_model)

        named_ranges: List[NamedRangeModel] = []
        try:
            for i in range(1, wb.Names.Count + 1):
                nm = wb.Names(i)
                named_ranges.append(
                    NamedRangeModel(
                        name=nm.Name,
                        worksheet=None,
                        value=str(nm.RefersTo),
                    )
                )
        except Exception:
            pass

        try:
            active_sheet_name = wb.ActiveSheet.Name if wb.ActiveSheet else (worksheets[0].name if worksheets else "")
        except Exception:
            active_sheet_name = worksheets[0].name if worksheets else ""

        return WorkbookModel(
            filename=filename,
            file_path=file_path,
            worksheets=worksheets,
            named_ranges=named_ranges,
            active_sheet=active_sheet_name,
        )

    @classmethod
    def _analyze_worksheet(cls, ws: Any, index: int) -> WorksheetModel:
        """Analyze a single worksheet COM object."""
        try:
            used_range = ws.UsedRange
            max_row = used_range.Row + used_range.Rows.Count - 1
            max_col = used_range.Column + used_range.Columns.Count - 1
            
            # Check if used range is genuinely empty (e.g. single blank cell A1)
            if used_range.Rows.Count == 1 and used_range.Columns.Count == 1 and used_range.Value is None:
                max_row = 0
                max_col = 0
        except Exception:
            max_row = 0
            max_col = 0

        # Count formulas via SpecialCells (xlCellTypeFormulas = -4123)
        formulas_count = 0
        if max_row > 0 and max_col > 0:
            try:
                f_cells = ws.Cells.SpecialCells(-4123)  # xlCellTypeFormulas
                formulas_count = f_cells.Count
            except Exception:
                formulas_count = 0

        tables: List[TableModel] = []
        processed_table_cells = set()

        # 1. Native Excel Tables (ListObjects)
        try:
            for i in range(1, ws.ListObjects.Count + 1):
                lo = ws.ListObjects(i)
                table_model = cls._analyze_native_table(ws, lo)
                tables.append(table_model)

                # Track cells in this table
                rng = lo.Range
                start_r = rng.Row
                start_c = rng.Column
                rows_c = rng.Rows.Count
                cols_c = rng.Columns.Count
                for r in range(start_r, start_r + rows_c):
                    for c in range(start_c, start_c + cols_c):
                        processed_table_cells.add((r, c))
        except Exception:
            pass

        # 2. Inferred Tabular Regions (if no native tables or additional non-empty regions exist)
        if not tables and max_row > 0 and max_col > 0:
            inferred_table = cls._infer_table_region(ws, max_row, max_col, processed_table_cells)
            if inferred_table:
                tables.append(inferred_table)

        # 3. Extract Charts (ChartObjects)
        charts: List[ChartModel] = []
        try:
            for i in range(1, ws.ChartObjects().Count + 1):
                co = ws.ChartObjects(i)
                try:
                    c_title = co.Chart.ChartTitle.Text if co.Chart.HasTitle else co.Name
                except Exception:
                    c_title = co.Name

                anchor_cell = "E2"
                try:
                    anchor_cell = co.TopLeftCell.Address.replace("$", "")
                except Exception:
                    pass

                charts.append(
                    ChartModel(
                        name=co.Name,
                        chart_type=str(co.Chart.ChartType) if hasattr(co.Chart, "ChartType") else "bar",
                        title=c_title,
                        worksheet=ws.Name,
                        cell_anchor=anchor_cell,
                    )
                )
        except Exception:
            pass

        return WorksheetModel(
            name=ws.Name,
            index=index,
            is_hidden=(ws.Visible != -1),  # xlSheetVisible = -1
            max_row=max_row,
            max_column=max_col,
            tables=tables,
            charts=charts,
            formulas_count=formulas_count,
        )

    @classmethod
    def _analyze_native_table(cls, ws: Any, lo: Any) -> TableModel:
        """Extract columns and info from native Excel ListObject."""
        name = lo.Name
        range_addr = lo.Range.Address.replace("$", "")
        row_count = lo.ListRows.Count if hasattr(lo, "ListRows") else 0

        columns: List[ColumnModel] = []
        try:
            for c_idx in range(1, lo.ListColumns.Count + 1):
                col = lo.ListColumns(c_idx)
                col_name = col.Name

                # Check formulas and samples
                has_formula = False
                sample_vals = []
                try:
                    body_rng = col.DataBodyRange
                    if body_rng is not None:
                        val_tuple = body_rng.Value
                        form_tuple = body_rng.Formula
                        if isinstance(form_tuple, tuple):
                            for r_row in form_tuple[:5]:
                                v = r_row[0] if isinstance(r_row, tuple) else r_row
                                if isinstance(v, str) and v.startswith("="):
                                    has_formula = True
                                    break
                        elif isinstance(form_tuple, str) and form_tuple.startswith("="):
                            has_formula = True

                        if isinstance(val_tuple, tuple):
                            for r_row in val_tuple[:5]:
                                v = r_row[0] if isinstance(r_row, tuple) else r_row
                                if v is not None:
                                    sample_vals.append(v)
                        elif val_tuple is not None:
                            sample_vals.append(val_tuple)
                except Exception:
                    pass

                columns.append(
                    ColumnModel(
                        name=col_name,
                        index=lo.Range.Column + c_idx - 1,
                        data_type="numeric" if sample_vals and isinstance(sample_vals[0], (int, float)) else "string",
                        has_formula=has_formula,
                        sample_values=sample_vals,
                    )
                )
        except Exception:
            pass

        return TableModel(
            name=name,
            worksheet=ws.Name,
            range=range_addr,
            is_native_table=True,
            columns=columns,
            row_count=row_count,
        )

    @classmethod
    def _infer_table_region(
        cls,
        ws: Any,
        max_row: int,
        max_col: int,
        processed_cells: set,
    ) -> Optional[TableModel]:
        """Infer table structure from raw cell grid when no native ListObject exists."""
        # Find first header row containing strings
        header_row = 1
        min_col = 1
        max_c = min(max_col, 50)
        max_r = min(max_row, 1000)

        # Read top slice
        try:
            slice_rng = ws.Range(ws.Cells(1, 1), ws.Cells(min(max_r, 10), max_c))
            vals = slice_rng.Value
        except Exception:
            return None

        if not vals:
            return None

        # Detect likely header row
        headers = []
        columns = []

        if isinstance(vals, tuple) and len(vals) > 0:
            first_row = vals[0]
            for c_idx in range(1, len(first_row) + 1):
                val = first_row[c_idx - 1]
                col_name = str(val).strip() if val is not None else f"Column_{c_idx}"
                headers.append(col_name)

                sample_vals = []
                has_formula = False
                for r_idx in range(1, min(len(vals), 6)):
                    cell_v = vals[r_idx][c_idx - 1]
                    if cell_v is not None:
                        sample_vals.append(cell_v)

                columns.append(
                    ColumnModel(
                        name=col_name,
                        index=c_idx,
                        data_type="numeric" if sample_vals and isinstance(sample_vals[0], (int, float)) else "string",
                        has_formula=has_formula,
                        sample_values=sample_vals,
                    )
                )

        range_str = f"{col_index_to_letter(min_col)}{header_row}:{col_index_to_letter(max_c)}{max_r}"

        return TableModel(
            name=f"{ws.Name}Table",
            worksheet=ws.Name,
            range=range_str,
            is_native_table=False,
            columns=columns,
            row_count=max(0, max_r - header_row),
        )

    @classmethod
    def generate_system_prompt(cls, model: WorkbookModel) -> str:
        """Generate formatted prompt context for the internal LLM."""
        compact = json.dumps(model.to_compact_dict(), indent=2)

        prompt = f"""SYSTEM INSTRUCTIONS: EXCEL AI COPILOT

You are an Excel AI Copilot. You analyze spreadsheet structures and propose precise workbook modifications directly on the user's active workbook.

CRITICAL RULES:
1. You NEVER edit Excel directly.
2. Refer to workbook elements SEMANTICALLY by Table name, Column name, or Worksheet name (NOT raw cell coordinates like B12 unless necessary).
3. Use semantic formula abstractions like `{{Profit}}/{{Revenue}}` or `SUM({{Revenue}})` instead of explicit cell coordinates.
4. When workbook interaction is required, include exactly ONE markdown block tagged with `excel-action`.

WORKBOOK CONTEXT:
```json
{compact}
```

FORMAT EXAMPLE:
I will add a Margin % column to the Sales table and update the dashboard chart.

```excel-action
{{
  "intent": "modify_workbook",
  "actions": [
    {{
      "action": "add_column",
      "table": "Sales",
      "column": "Margin %"
    }},
    {{
      "action": "insert_formula",
      "table": "Sales",
      "column": "Margin %",
      "formula": "{{Profit}}/{{Revenue}}"
    }}
  ]
}}
```
"""
        return prompt
