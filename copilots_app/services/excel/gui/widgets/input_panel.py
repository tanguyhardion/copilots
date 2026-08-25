"""Input Panel Widget: Textarea for pasting LLM responses and triggering action parsing."""

from typing import Optional, Callable
from PySide6.QtWidgets import (
    QWidget,
    QVBoxLayout,
    QHBoxLayout,
    QGroupBox,
    QPlainTextEdit,
    QPushButton,
    QLabel,
)
from PySide6.QtCore import Signal
from copilots_app.services.excel.protocol.action_parser import ActionParser
from copilots_app.services.excel.models.protocol import ActionProtocol, ActionIntent


class InputPanelWidget(QWidget):
    """Textarea for pasting LLM response and parsing Action Protocol blocks."""

    protocol_parsed = Signal(ActionProtocol)

    def __init__(self, parent: Optional[QWidget] = None):
        super().__init__(parent)
        self._init_ui()

    def _init_ui(self):
        layout = QVBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)

        group = QGroupBox("LLM Response Input")
        group_layout = QVBoxLayout(group)

        info_lbl = QLabel("Paste the LLM's response below (supports full text with embedded excel-action blocks):")
        group_layout.addWidget(info_lbl)

        self.input_edit = QPlainTextEdit()
        self.input_edit.setPlaceholderText(
            "Example LLM Response:\n\n"
            "I will add a Margin % column and update the dashboard chart.\n\n"
            "```excel-action\n"
            "{\n"
            '  "intent": "modify_workbook",\n'
            '  "actions": [\n'
            '    {"action": "add_column", "table": "Sales", "column": "Margin %"},\n'
            '    {"action": "insert_formula", "table": "Sales", "column": "Margin %", "formula": "{Profit}/{Revenue}"}\n'
            "  ]\n"
            "}\n"
            "```"
        )
        self.input_edit.textChanged.connect(self._on_text_changed)
        group_layout.addWidget(self.input_edit)

        # Status Bar & Action Button
        btn_layout = QHBoxLayout()
        self.status_lbl = QLabel("Status: Ready for input")
        self.status_lbl.setStyleSheet("color: #94A3B8; font-weight: bold;")

        self.parse_btn = QPushButton("Parse & Validate Actions")
        self.parse_btn.clicked.connect(self.parse_input)

        btn_layout.addWidget(self.status_lbl)
        btn_layout.addStretch()
        btn_layout.addWidget(self.parse_btn)

        group_layout.addLayout(btn_layout)
        layout.addWidget(group)

    def _on_text_changed(self):
        text = self.input_edit.toPlainText()
        if "excel-action" in text or '"intent"' in text:
            self.status_lbl.setText("Status: Protocol block detected! Click 'Parse & Validate'")
            self.status_lbl.setStyleSheet("color: #38BDF8; font-weight: bold;")
        elif text.strip():
            self.status_lbl.setText("Status: Conversation text detected (No action block)")
            self.status_lbl.setStyleSheet("color: #F59E0B; font-weight: bold;")
        else:
            self.status_lbl.setText("Status: Ready for input")
            self.status_lbl.setStyleSheet("color: #94A3B8; font-weight: bold;")

    def parse_input(self):
        text = self.input_edit.toPlainText()
        protocol = ActionParser.parse_response(text)
        self.protocol_parsed.emit(protocol)

    def clear(self):
        self.input_edit.clear()
