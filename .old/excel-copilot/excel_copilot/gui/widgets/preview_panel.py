"""Preview Panel Widget: Shows pre-execution validation feedback, action diff preview, and approval controls."""

from typing import Optional
from PySide6.QtWidgets import (
    QWidget,
    QVBoxLayout,
    QHBoxLayout,
    QGroupBox,
    QListWidget,
    QListWidgetItem,
    QLabel,
    QPushButton,
    QCheckBox,
)
from PySide6.QtCore import Signal, Qt
from excel_copilot.models.protocol import ActionProtocol, ValidationResult
from excel_copilot.models.semantic import WorkbookModel
from excel_copilot.validator.action_validator import ActionValidator


class PreviewPanelWidget(QWidget):
    """Displays action previews, validation state, and execution approval button."""

    execute_requested = Signal()

    def __init__(self, parent: Optional[QWidget] = None):
        super().__init__(parent)
        self.current_protocol: Optional[ActionProtocol] = None
        self.current_validation: Optional[ValidationResult] = None
        self._init_ui()

    def _init_ui(self):
        layout = QVBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)

        group = QGroupBox("Action Validation & Execution Diff Preview")
        group_layout = QVBoxLayout(group)

        # Status Summary Header
        self.summary_lbl = QLabel("No action protocol parsed yet.")
        self.summary_lbl.setStyleSheet("font-weight: bold; font-size: 14px;")
        group_layout.addWidget(self.summary_lbl)

        # List Widget for Action Items / Diff Preview
        self.preview_list = QListWidget()
        group_layout.addWidget(self.preview_list)

        # Confirm & Approve Execution Bar
        controls_layout = QHBoxLayout()
        self.confirm_cb = QCheckBox("I approve these Excel modifications")
        self.confirm_cb.setEnabled(False)
        self.confirm_cb.toggled.connect(self._on_confirm_toggled)

        self.execute_btn = QPushButton("Approve & Execute Actions")
        self.execute_btn.setObjectName("SuccessButton")
        self.execute_btn.setEnabled(False)
        self.execute_btn.clicked.connect(self._on_execute_clicked)

        controls_layout.addWidget(self.confirm_cb)
        controls_layout.addStretch()
        controls_layout.addWidget(self.execute_btn)

        group_layout.addLayout(controls_layout)
        layout.addWidget(group)

    def set_protocol_and_validate(self, protocol: ActionProtocol, model: Optional[WorkbookModel]):
        """Run validation engine and display results."""
        self.current_protocol = protocol
        self.current_validation = ActionValidator.validate(protocol, model)
        self.preview_list.clear()

        val = self.current_validation

        if not val.is_valid:
            self.summary_lbl.setText("Validation Failed! Fix blocking errors before executing:")
            self.summary_lbl.setStyleSheet("color: #EF4444; font-weight: bold; font-size: 14px;")

            for err in val.errors:
                item = QListWidgetItem(f"❌ ERROR: {err}")
                item.setForeground(Qt.red)
                self.preview_list.addItem(item)

            for warn in val.warnings:
                item = QListWidgetItem(f"⚠️ WARNING: {warn}")
                item.setForeground(Qt.yellow)
                self.preview_list.addItem(item)

            self.confirm_cb.setEnabled(False)
            self.confirm_cb.setChecked(False)
            self.execute_btn.setEnabled(False)

        elif not val.action_previews or "No actions" in val.action_previews[0]:
            self.summary_lbl.setText("Conversation Response (No workbook actions to execute).")
            self.summary_lbl.setStyleSheet("color: #94A3B8; font-weight: bold; font-size: 14px;")
            self.preview_list.addItem("No Excel modifications requested by LLM.")
            self.confirm_cb.setEnabled(False)
            self.confirm_cb.setChecked(False)
            self.execute_btn.setEnabled(False)

        else:
            self.summary_lbl.setText(f"Validation Passed! {len(val.action_previews)} action(s) ready to execute:")
            self.summary_lbl.setStyleSheet("color: #10B981; font-weight: bold; font-size: 14px;")

            for idx, prev in enumerate(val.action_previews, start=1):
                item = QListWidgetItem(f"⚡ Step {idx}: {prev}")
                item.setForeground(Qt.cyan)
                self.preview_list.addItem(item)

            for warn in val.warnings:
                item = QListWidgetItem(f"⚠️ WARNING: {warn}")
                item.setForeground(Qt.yellow)
                self.preview_list.addItem(item)

            self.confirm_cb.setEnabled(True)
            self.confirm_cb.setChecked(False)
            self.execute_btn.setEnabled(False)

    def _on_confirm_toggled(self, checked: bool):
        if self.current_validation and self.current_validation.is_valid:
            self.execute_btn.setEnabled(checked)

    def _on_execute_clicked(self):
        self.execute_requested.emit()
