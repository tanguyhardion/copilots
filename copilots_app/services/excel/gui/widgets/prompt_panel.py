"""Prompt Panel Widget: Generates and formats copyable system prompt and workbook context for LLM."""

from typing import Optional
from PySide6.QtWidgets import (
    QWidget,
    QVBoxLayout,
    QHBoxLayout,
    QGroupBox,
    QPlainTextEdit,
    QPushButton,
    QLabel,
    QApplication,
    QMessageBox,
)
from copilots_app.services.excel.models.semantic import WorkbookModel
from copilots_app.services.excel.analyzer.workbook_analyzer import WorkbookAnalyzer


class PromptPanelWidget(QWidget):
    """Displays generated LLM System Prompt and copy controls."""

    def __init__(self, parent: Optional[QWidget] = None):
        super().__init__(parent)
        self._init_ui()

    def _init_ui(self):
        layout = QVBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)

        group = QGroupBox("LLM Prompt & Workbook Context Generator")
        group_layout = QVBoxLayout(group)

        # Header Info
        info_lbl = QLabel(
            "Paste this System Prompt and Semantic Model into your LLM conversation to provide workbook context:"
        )
        info_lbl.setWordWrap(True)
        group_layout.addWidget(info_lbl)

        # Prompt View
        self.prompt_edit = QPlainTextEdit()
        self.prompt_edit.setReadOnly(True)
        group_layout.addWidget(self.prompt_edit)

        # Copy Action Button
        btn_layout = QHBoxLayout()
        self.copy_btn = QPushButton("Copy LLM Context to Clipboard")
        self.copy_btn.setObjectName("SuccessButton")
        self.copy_btn.clicked.connect(self._on_copy_clicked)
        btn_layout.addStretch()
        btn_layout.addWidget(self.copy_btn)

        group_layout.addLayout(btn_layout)
        layout.addWidget(group)

    def update_prompt(self, model: Optional[WorkbookModel]):
        """Generate and display system prompt string."""
        if not model:
            self.prompt_edit.setPlainText("No workbook loaded. Open an Excel file to generate LLM context.")
            return

        prompt_str = WorkbookAnalyzer.generate_system_prompt(model)
        self.prompt_edit.setPlainText(prompt_str)

    def _on_copy_clicked(self):
        text = self.prompt_edit.toPlainText()
        if text and "No workbook loaded" not in text:
            clipboard = QApplication.clipboard()
            clipboard.setText(text)
            QMessageBox.information(
                self, "Copied", "LLM System Prompt and Semantic Model copied to clipboard!"
            )
