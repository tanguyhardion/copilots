"""Main Application Window for Excel AI Copilot Desktop App."""

import sys
import os
from typing import Optional

from PySide6.QtWidgets import (
    QApplication,
    QMainWindow,
    QWidget,
    QVBoxLayout,
    QHBoxLayout,
    QSplitter,
    QTabWidget,
    QToolBar,
    QFileDialog,
    QMessageBox,
    QStatusBar,
    QLabel,
)
from PySide6.QtGui import QAction, QIcon
from PySide6.QtCore import Qt

from excel_copilot.models.semantic import WorkbookModel
from excel_copilot.models.protocol import ActionProtocol
from excel_copilot.analyzer.workbook_analyzer import WorkbookAnalyzer
from excel_copilot.executor.executor_main import ActionExecutor
from excel_copilot.gui.styles import DARK_THEME
from excel_copilot.gui.widgets import (
    WorkbookExplorerWidget,
    PromptPanelWidget,
    InputPanelWidget,
    PreviewPanelWidget,
    LogPanelWidget,
)


class MainWindow(QMainWindow):
    """Main window for Excel AI Copilot desktop application."""

    def __init__(self, sample_file: Optional[str] = None):
        super().__init__()
        self.setWindowTitle("Excel AI Copilot — Air-Gapped Desktop Suite")
        self.resize(1280, 800)

        self.current_file_path: Optional[str] = None
        self.current_model: Optional[WorkbookModel] = None
        self.current_protocol: Optional[ActionProtocol] = None
        self.executor = ActionExecutor()

        self._init_ui()

        if sample_file and os.path.exists(sample_file):
            self.load_workbook(sample_file)

    def _init_ui(self):
        self.setStyleSheet(DARK_THEME)

        # 1. Setup Toolbar
        toolbar = QToolBar("Main Controls")
        self.addToolBar(toolbar)

        open_act = QAction("📂 Open Workbook", self)
        open_act.triggered.connect(self._on_open_file)
        toolbar.addAction(open_act)

        save_act = QAction("💾 Save Workbook", self)
        save_act.triggered.connect(self._on_save_file)
        toolbar.addAction(save_act)

        save_as_act = QAction("💾 Save As...", self)
        save_as_act.triggered.connect(self._on_save_as_file)
        toolbar.addAction(save_as_act)

        toolbar.addSeparator()

        backup_act = QAction("🛡️ Backups", self)
        backup_act.triggered.connect(self._on_view_backups)
        toolbar.addAction(backup_act)

        # 2. Main Central Splitter
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        main_layout = QHBoxLayout(central_widget)
        main_layout.setContentsMargins(8, 8, 8, 8)

        splitter = QSplitter(Qt.Horizontal)

        # Left Tabs (Workbook & Context)
        self.left_tabs = QTabWidget()
        self.explorer_widget = WorkbookExplorerWidget()
        self.prompt_widget = PromptPanelWidget()

        self.left_tabs.addTab(self.explorer_widget, "Workbook Explorer")
        self.left_tabs.addTab(self.prompt_widget, "System Prompt Context")
        splitter.addWidget(self.left_tabs)

        # Right Tabs (Input, Preview, Logs)
        self.right_tabs = QTabWidget()
        self.input_widget = InputPanelWidget()
        self.preview_widget = PreviewPanelWidget()
        self.log_widget = LogPanelWidget()

        self.right_tabs.addTab(self.input_widget, "Command Input")
        self.right_tabs.addTab(self.preview_widget, "Validation & Action Preview")
        self.right_tabs.addTab(self.log_widget, "Execution Report")
        splitter.addWidget(self.right_tabs)

        splitter.setSizes([500, 780])
        main_layout.addWidget(splitter)

        # Connect Signals
        self.input_widget.protocol_parsed.connect(self._on_protocol_parsed)
        self.preview_widget.execute_requested.connect(self._on_execute_requested)

        # 3. Status Bar
        self.statusBar = QStatusBar()
        self.setStatusBar(self.statusBar)
        self.status_file_lbl = QLabel("No workbook loaded")
        self.statusBar.addPermanentWidget(self.status_file_lbl)

    def _on_open_file(self):
        file_path, _ = QFileDialog.getOpenFileName(
            self, "Open Excel Workbook", "", "Excel Files (*.xlsx *.xlsm)"
        )
        if file_path:
            self.load_workbook(file_path)

    def load_workbook(self, file_path: str):
        """Analyze and load target Excel workbook."""
        try:
            self.current_file_path = os.path.abspath(file_path)
            self.current_model = WorkbookAnalyzer.analyze(self.current_file_path)

            self.explorer_widget.load_workbook_model(self.current_model)
            self.prompt_widget.update_prompt(self.current_model)

            self.status_file_lbl.setText(f"Active File: {self.current_file_path}")
            self.statusBar.showMessage(f"Successfully loaded '{os.path.basename(file_path)}'", 4000)

        except Exception as err:
            QMessageBox.critical(self, "Error Opening File", f"Failed to load workbook:\n{str(err)}")

    def _on_protocol_parsed(self, protocol: ActionProtocol):
        """Handle parsed protocol from input panel."""
        self.current_protocol = protocol
        self.preview_widget.set_protocol_and_validate(protocol, self.current_model)
        self.right_tabs.setCurrentWidget(self.preview_widget)

    def _on_execute_requested(self):
        """Execute validated action protocol."""
        if not self.current_protocol or not self.current_file_path:
            QMessageBox.warning(self, "Execution Error", "No active workbook or action protocol available.")
            return

        result, updated_model = self.executor.execute(
            protocol=self.current_protocol,
            file_path=self.current_file_path,
            model=self.current_model,
            save_file=True,
        )

        if updated_model:
            self.current_model = updated_model
            self.explorer_widget.load_workbook_model(self.current_model)
            self.prompt_widget.update_prompt(self.current_model)

        self.log_widget.display_execution_result(result)
        self.right_tabs.setCurrentWidget(self.log_widget)

        self.statusBar.showMessage(
            f"Execution finished with status '{result.status.value}'. Executed {result.actions_executed} action(s).",
            6000,
        )

    def _on_save_file(self):
        if self.current_file_path:
            QMessageBox.information(self, "Saved", f"Workbook saved to '{self.current_file_path}'")

    def _on_save_as_file(self):
        if not self.current_file_path:
            return
        dest_path, _ = QFileDialog.getSaveFileName(
            self, "Save Workbook As", "", "Excel Files (*.xlsx)"
        )
        if dest_path:
            import shutil
            shutil.copy2(self.current_file_path, dest_path)
            self.load_workbook(dest_path)

    def _on_view_backups(self):
        backups = self.executor.backup_manager.list_backups()
        if not backups:
            QMessageBox.information(self, "Backups", "No backups created yet.")
            return

        msg = f"Available Backups ({len(backups)}):\n\n"
        for b in backups[:5]:
            msg += f"• {b['filename']} ({b['created_at']})\n"

        QMessageBox.information(self, "Backups Manager", msg)


def main():
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
