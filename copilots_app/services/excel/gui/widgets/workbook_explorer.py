"""Workbook Explorer Widget: Visual tree view of worksheets, tables, columns, and charts."""

from typing import Optional
from PySide6.QtWidgets import (
    QWidget,
    QVBoxLayout,
    QTreeWidget,
    QTreeWidgetItem,
    QGroupBox,
    QTableWidget,
    QTableWidgetItem,
    QLabel,
    QSplitter,
)
from PySide6.QtCore import Qt
from copilots_app.services.excel.models.semantic import WorkbookModel, WorksheetModel


class WorkbookExplorerWidget(QWidget):
    """Displays structured tree of open workbook and sample data preview."""

    def __init__(self, parent: Optional[QWidget] = None):
        super().__init__(parent)
        self._init_ui()

    def _init_ui(self):
        layout = QVBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)

        splitter = QSplitter(Qt.Vertical)

        # 1. Tree View Group
        tree_group = QGroupBox("Workbook Objects Explorer")
        tree_layout = QVBoxLayout(tree_group)

        self.tree_widget = QTreeWidget()
        self.tree_widget.setHeaderLabels(["Object Name", "Type / Range", "Details"])
        tree_layout.addWidget(self.tree_widget)
        splitter.addWidget(tree_group)

        # 2. Sample Data Table Group
        preview_group = QGroupBox("Worksheet Data Sample Preview")
        preview_layout = QVBoxLayout(preview_group)

        self.sample_table = QTableWidget()
        preview_layout.addWidget(self.sample_table)
        splitter.addWidget(preview_group)

        splitter.setSizes([300, 200])
        layout.addWidget(splitter)

    def load_workbook_model(self, model: Optional[WorkbookModel]):
        """Populate tree view with WorkbookModel objects."""
        self.tree_widget.clear()
        self.sample_table.clear()
        self.sample_table.setRowCount(0)
        self.sample_table.setColumnCount(0)

        if not model:
            return

        # Root Item
        wb_item = QTreeWidgetItem(self.tree_widget, [model.filename, "Workbook", f"{len(model.worksheets)} sheet(s)"])
        wb_item.setExpanded(True)

        for ws in model.worksheets:
            ws_item = QTreeWidgetItem(wb_item, [ws.name, "Worksheet", f"{ws.max_row} rows, {ws.max_column} cols"])
            ws_item.setExpanded(True)

            # Tables
            if ws.tables:
                tables_group = QTreeWidgetItem(ws_item, ["Tables", "Group", f"{len(ws.tables)} table(s)"])
                tables_group.setExpanded(True)
                for tbl in ws.tables:
                    tbl_item = QTreeWidgetItem(tables_group, [tbl.name, "Table", tbl.range])
                    tbl_item.setExpanded(True)
                    for col in tbl.columns:
                        sample_str = f"Sample: {col.sample_values[:2]}" if col.sample_values else ""
                        QTreeWidgetItem(tbl_item, [col.name, f"Column ({col.data_type})", sample_str])

            # Charts
            if ws.charts:
                charts_group = QTreeWidgetItem(ws_item, ["Charts", "Group", f"{len(ws.charts)} chart(s)"])
                for chart in ws.charts:
                    QTreeWidgetItem(charts_group, [chart.name, f"Chart ({chart.chart_type})", f"Anchor {chart.cell_anchor}"])

        # Populate sample table from first table in active sheet
        if model.worksheets:
            first_ws = model.worksheets[0]
            if first_ws.tables:
                tbl = first_ws.tables[0]
                cols = tbl.columns
                self.sample_table.setColumnCount(len(cols))
                self.sample_table.setHorizontalHeaderLabels([c.name for c in cols])

                # Max sample rows
                max_samples = max((len(c.sample_values) for c in cols), default=0)
                self.sample_table.setRowCount(max_samples)

                for c_idx, c in enumerate(cols):
                    for r_idx, val in enumerate(c.sample_values):
                        item = QTableWidgetItem(str(val))
                        self.sample_table.setItem(r_idx, c_idx, item)
