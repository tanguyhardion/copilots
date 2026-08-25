"""Log Panel Widget: Displays execution logs, backup details, and copyable Execution Result JSON."""

import json
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
from excel_copilot.models.protocol import ExecutionResult


class LogPanelWidget(QWidget):
    """Displays execution history and formatted Execution Result for copying back to LLM."""

    def __init__(self, parent: Optional[QWidget] = None):
        super().__init__(parent)
        self._init_ui()

    def _init_ui(self):
        layout = QVBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)

        group = QGroupBox("Execution Log & Structured Result Exporter")
        group_layout = QVBoxLayout(group)

        info_lbl = QLabel("Copy this Execution Result JSON back into your LLM conversation to complete the feedback loop:")
        group_layout.addWidget(info_lbl)

        self.log_edit = QPlainTextEdit()
        self.log_edit.setReadOnly(True)
        group_layout.addWidget(self.log_edit)

        # Action Buttons
        btn_layout = QHBoxLayout()
        self.copy_result_btn = QPushButton("Copy Execution Result to Clipboard")
        self.copy_result_btn.setObjectName("SuccessButton")
        self.copy_result_btn.clicked.connect(self._on_copy_clicked)

        btn_layout.addStretch()
        btn_layout.addWidget(self.copy_result_btn)

        group_layout.addLayout(btn_layout)
        layout.addWidget(group)

    def display_execution_result(self, result: ExecutionResult):
        """Format and display ExecutionResult model."""
        res_dict = result.dict(exclude_none=True)
        res_json = json.dumps(res_dict, indent=2)

        log_text = f"=== EXECUTION REPORT ===\n"
        log_text += f"Status: {result.status.value.upper()}\n"
        log_text += f"Actions Executed: {result.actions_executed}\n"
        log_text += f"Objects Modified: {', '.join(result.objects_modified) if result.objects_modified else 'None'}\n"

        if result.errors:
            log_text += f"\nERRORS ({len(result.errors)}):\n"
            for err in result.errors:
                log_text += f" - {err}\n"

        if result.warnings:
            log_text += f"\nWARNINGS ({len(result.warnings)}):\n"
            for warn in result.warnings:
                log_text += f" - {warn}\n"

        log_text += f"\n=== STRUCTURED EXECUTION RESULT (Copy to LLM) ===\n"
        log_text += f"```json\n{res_json}\n```"

        self.log_edit.setPlainText(log_text)

    def append_log(self, text: str):
        self.log_edit.appendPlainText(text)

    def _on_copy_clicked(self):
        text = self.log_edit.toPlainText()
        if "=== STRUCTURED EXECUTION RESULT" in text:
            # Extract JSON block
            parts = text.split("```json\n")
            if len(parts) > 1:
                json_block = parts[1].split("\n```")[0]
                clipboard = QApplication.clipboard()
                clipboard.setText(json_block)
                QMessageBox.information(
                    self, "Copied", "Execution Result JSON copied to clipboard! Paste this into your LLM conversation."
                )
                return

        # Fallback copy full log
        clipboard = QApplication.clipboard()
        clipboard.setText(text)
        QMessageBox.information(self, "Copied", "Log output copied to clipboard!")
