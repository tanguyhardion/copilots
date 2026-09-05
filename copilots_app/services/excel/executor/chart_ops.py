"""Chart operations: create, update, delete Excel charts using pywin32 COM."""

from typing import Optional, Any


class ChartOps:
    """Executes chart modifications on active Excel Workbook COM object."""

    @classmethod
    def _get_sheet(cls, wb: Any, sheet_name: Optional[str]) -> Any:
        if not sheet_name:
            return wb.ActiveSheet
        for i in range(1, wb.Worksheets.Count + 1):
            if wb.Worksheets(i).Name.lower() == sheet_name.lower():
                return wb.Worksheets(i)
        return wb.ActiveSheet

    @classmethod
    def execute_create_chart(
        cls,
        wb: Any,
        sheet_name: str,
        chart_title: str = "Chart",
        chart_type: str = "bar",
        data_range: Optional[str] = None,
        anchor: str = "E2",
    ) -> str:
        """Create and embed a new Excel chart."""
        ws = cls._get_sheet(wb, sheet_name)

        c_type_str = str(chart_type).lower().strip()
        # Excel ChartType Enums:
        # xlColumnClustered = 51, xlBarClustered = 57, xlLine = 4, xlPie = 5
        if "line" in c_type_str:
            xl_chart_type = 4
        elif "pie" in c_type_str:
            xl_chart_type = 5
        elif "bar" in c_type_str and "col" not in c_type_str:
            xl_chart_type = 57
        else:
            xl_chart_type = 51  # xlColumnClustered

        anchor_cell = ws.Range(anchor)
        left = anchor_cell.Left
        top = anchor_cell.Top
        width = 380
        height = 240

        chart_obj = ws.ChartObjects().Add(left, top, width, height)
        chart = chart_obj.Chart
        chart.ChartType = xl_chart_type

        # Source data
        if data_range and ":" in data_range:
            src_rng = ws.Range(data_range)
        else:
            src_rng = ws.UsedRange

        chart.SetSourceData(src_rng)
        chart.HasTitle = True
        chart.ChartTitle.Text = chart_title

        return f"Created {c_type_str.capitalize()} chart '{chart_title}' anchored at {anchor} in '{ws.Name}'."

    @classmethod
    def execute_delete_chart(cls, wb: Any, sheet_name: str, chart_name: str) -> str:
        """Remove a chart by title or name."""
        ws = cls._get_sheet(wb, sheet_name)

        deleted = 0
        try:
            for i in range(ws.ChartObjects().Count, 0, -1):
                co = ws.ChartObjects(i)
                title = co.Chart.ChartTitle.Text if co.Chart.HasTitle else co.Name
                if chart_name.lower() in title.lower() or chart_name.lower() in co.Name.lower():
                    co.Delete()
                    deleted += 1
        except Exception:
            pass

        return f"Deleted {deleted} chart(s) matching '{chart_name}' from '{ws.Name}'."
