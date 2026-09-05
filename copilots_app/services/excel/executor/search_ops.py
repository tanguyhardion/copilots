"""Search operations: Query sheets, tables, columns, formulas, and cell text using pywin32 COM."""

from typing import List, Dict, Any
from copilots_app.services.excel.models.semantic import WorkbookModel


class SearchOps:
    """Executes read-only query operations across active Excel workbook."""

    @classmethod
    def find_sheet(cls, model: WorkbookModel, query: str) -> List[Dict[str, Any]]:
        """Find worksheets matching query string."""
        results = []
        q = query.lower().strip()
        for ws in model.worksheets:
            if q in ws.name.lower():
                results.append({
                    "sheet_name": ws.name,
                    "tables_count": len(ws.tables),
                    "charts_count": len(ws.charts),
                    "max_row": ws.max_row,
                })
        return results

    @classmethod
    def find_table(cls, model: WorkbookModel, query: str) -> List[Dict[str, Any]]:
        """Find tables matching query string."""
        results = []
        q = query.lower().strip()
        for ws in model.worksheets:
            for tbl in ws.tables:
                if q in tbl.name.lower() or q in ws.name.lower():
                    results.append({
                        "table_name": tbl.name,
                        "worksheet": ws.name,
                        "range": tbl.range,
                        "columns": [c.name for c in tbl.columns],
                        "row_count": tbl.row_count,
                    })
        return results

    @classmethod
    def find_column(cls, model: WorkbookModel, query: str) -> List[Dict[str, Any]]:
        """Find columns matching query string."""
        results = []
        q = query.lower().strip()
        for ws in model.worksheets:
            for tbl in ws.tables:
                for col in tbl.columns:
                    if q in col.name.lower():
                        results.append({
                            "column_name": col.name,
                            "table_name": tbl.name,
                            "worksheet": ws.name,
                            "data_type": col.data_type,
                            "has_formula": col.has_formula,
                        })
        return results

    @classmethod
    def search_text(cls, wb: Any, query: str) -> List[Dict[str, Any]]:
        """Search text across all cells in active workbook using Excel COM Range.Find."""
        results = []
        q = str(query).strip()
        if not q:
            return results

        # xlValues = -4163, xlPart = 2
        for i in range(1, wb.Worksheets.Count + 1):
            ws = wb.Worksheets(i)
            try:
                first_found = ws.UsedRange.Find(What=q, LookIn=-4163, LookAt=2)
                if first_found:
                    first_addr = first_found.Address
                    curr = first_found
                    while True:
                        results.append({
                            "sheet": ws.Name,
                            "coordinate": curr.Address.replace("$", ""),
                            "value": str(curr.Value),
                        })
                        if len(results) >= 50:
                            return results
                        curr = ws.UsedRange.FindNext(curr)
                        if not curr or curr.Address == first_addr:
                            break
            except Exception:
                pass

        return results
