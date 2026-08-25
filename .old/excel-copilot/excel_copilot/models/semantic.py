"""Semantic data models representing Excel workbook structure."""

from typing import List, Optional, Any, Dict
from pydantic import BaseModel, Field


class ColumnModel(BaseModel):
    """Represents a column in a table or worksheet region."""

    name: str = Field(..., description="Column header name")
    index: Optional[int] = Field(None, description="1-based column index")
    data_type: str = Field("string", description="Inferred data type (numeric, date, string, boolean, formula)")
    has_formula: bool = Field(False, description="Whether column contains formulas")
    sample_values: List[Any] = Field(default_factory=list, description="Sample values from top rows")


class TableModel(BaseModel):
    """Represents an Excel Table or an inferred tabular data region."""

    name: str = Field(..., description="Table name (e.g. SalesData)")
    worksheet: str = Field(..., description="Parent worksheet name")
    range: str = Field(..., description="Cell range (e.g. A1:D50)")
    is_native_table: bool = Field(True, description="True if official Excel Table, False if inferred region")
    columns: List[ColumnModel] = Field(default_factory=list, description="List of columns in table")
    row_count: int = Field(0, description="Total data rows in table")


class ChartModel(BaseModel):
    """Represents a chart object in a worksheet."""

    name: str = Field(..., description="Chart name or title")
    chart_type: str = Field("bar", description="Type of chart (bar, line, pie, column, scatter)")
    title: str = Field("", description="Chart title")
    worksheet: str = Field(..., description="Parent worksheet name")
    cell_anchor: str = Field("E2", description="Anchor cell location")
    data_range: Optional[str] = Field(None, description="Source data reference range")


class NamedRangeModel(BaseModel):
    """Represents a defined named range in the workbook."""

    name: str = Field(..., description="Name of defined range")
    worksheet: Optional[str] = Field(None, description="Target worksheet if scope is local")
    value: str = Field(..., description="Formula or reference string (e.g. Sheet1!$A$1:$B$10)")


class WorksheetModel(BaseModel):
    """Represents an individual worksheet inside the workbook."""

    name: str = Field(..., description="Worksheet tab name")
    index: int = Field(0, description="0-based tab index")
    is_hidden: bool = Field(False, description="Hidden tab status")
    max_row: int = Field(0, description="Highest populated row")
    max_column: int = Field(0, description="Highest populated column")
    tables: List[TableModel] = Field(default_factory=list, description="Tables within sheet")
    charts: List[ChartModel] = Field(default_factory=list, description="Charts within sheet")
    unbound_columns: List[ColumnModel] = Field(default_factory=list, description="Columns outside tables")
    formulas_count: int = Field(0, description="Total formulas in sheet")


class WorkbookModel(BaseModel):
    """Complete semantic model of an Excel workbook for LLM reasoning."""

    filename: str = Field(..., description="Workbook file name")
    file_path: Optional[str] = Field(None, description="Absolute file path on disk")
    worksheets: List[WorksheetModel] = Field(default_factory=list, description="List of worksheets")
    named_ranges: List[NamedRangeModel] = Field(default_factory=list, description="Defined named ranges")
    active_sheet: str = Field("", description="Currently active worksheet name")

    def get_sheet(self, sheet_name: str) -> Optional[WorksheetModel]:
        """Find worksheet by name (case-insensitive)."""
        for ws in self.worksheets:
            if ws.name.strip().lower() == sheet_name.strip().lower():
                return ws
        return None

    def get_table(self, table_name: str) -> Optional[TableModel]:
        """Find table across all sheets by name (case-insensitive)."""
        for ws in self.worksheets:
            for tbl in ws.tables:
                if tbl.name.strip().lower() == table_name.strip().lower():
                    return tbl
        return None

    def to_compact_dict(self) -> Dict[str, Any]:
        """Generate compact dictionary for LLM context prompt."""
        return {
            "filename": self.filename,
            "worksheets": [
                {
                    "name": ws.name,
                    "tables": [
                        {
                            "name": tbl.name,
                            "range": tbl.range,
                            "is_native": tbl.is_native_table,
                            "columns": [col.name for col in tbl.columns],
                            "row_count": tbl.row_count,
                        }
                        for tbl in ws.tables
                    ],
                    "charts": [c.name for c in ws.charts],
                }
                for ws in self.worksheets
            ],
            "named_ranges": [{"name": nr.name, "value": nr.value} for nr in self.named_ranges],
        }
