"""
Prompt Management Dialog in PySide6: View, Edit, Copy, Save Custom, and Reset System Prompts.
"""

from typing import Optional, Callable
from PySide6.QtWidgets import (
    QDialog,
    QWidget,
    QVBoxLayout,
    QHBoxLayout,
    QLabel,
    QPushButton,
    QPlainTextEdit,
    QFrame,
    QApplication,
)
from PySide6.QtCore import Qt
from PySide6.QtGui import QFont

from copilots_app.core.theme import AppPalette
from copilots_app.core.prompt_manager import PromptManager


class PromptDialog(QDialog):
    """Modern modal dialog for editing, copying, and resetting LLM System Prompts."""

    def __init__(
        self,
        copilot_key: str,
        on_status_change: Optional[Callable[[str, str], None]] = None,
        parent: Optional[QWidget] = None,
    ):
        super().__init__(parent)
        self.copilot_key = copilot_key
        self.on_status_change = on_status_change
        self.pm = PromptManager()

        self.setWindowTitle(self.pm.PROMPT_TITLES.get(copilot_key, f"{copilot_key.capitalize()} System Prompt"))
        self.resize(880, 580)
        self.setStyleSheet(f"""
            QDialog {{
                background-color: {AppPalette.BG_DARK};
            }}
        """)

        layout = QVBoxLayout(self)
        layout.setContentsMargins(20, 20, 20, 20)
        layout.setSpacing(14)

        # Header Info Row
        header_row = QHBoxLayout()
        header_row.setSpacing(10)

        title_label = QLabel(self.pm.PROMPT_TITLES.get(copilot_key, f"{copilot_key.capitalize()} System Prompt"))
        title_label.setFont(QFont("Segoe UI", 12, QFont.Bold))
        title_label.setStyleSheet(f"color: {AppPalette.TEXT_PRIMARY};")
        header_row.addWidget(title_label)

        self.is_custom = self.pm.is_customized(copilot_key)
        self.badge_label = QLabel("User Custom Override" if self.is_custom else "Default Bundled")
        self.badge_label.setFont(QFont("Segoe UI", 8, QFont.Bold))
        self._update_badge_style()
        header_row.addWidget(self.badge_label)

        header_row.addStretch()

        open_folder_btn = QPushButton("Open Prompts Folder")
        open_folder_btn.setFont(QFont("Segoe UI", 9))
        open_folder_btn.setStyleSheet(f"""
            QPushButton {{
                background-color: {AppPalette.BG_CARD};
                color: {AppPalette.TEXT_SECONDARY};
                border: 1px solid {AppPalette.BORDER_COLOR};
                border-radius: 5px;
                padding: 5px 12px;
            }}
            QPushButton:hover {{
                background-color: {AppPalette.BG_CARD_HOVER};
                color: {AppPalette.TEXT_PRIMARY};
            }}
        """)
        open_folder_btn.clicked.connect(lambda: self.pm.open_prompts_directory())
        header_row.addWidget(open_folder_btn)

        layout.addLayout(header_row)

        desc_label = QLabel("This system prompt instructs the LLM how to format responses and DSL code for this copilot.")
        desc_label.setFont(QFont("Segoe UI", 9))
        desc_label.setStyleSheet(f"color: {AppPalette.TEXT_MUTED};")
        layout.addWidget(desc_label)

        # Prompt Text Editor
        self.editor = QPlainTextEdit()
        self.editor.setPlainText(self.pm.get_prompt(copilot_key))
        self.editor.setFont(QFont("Consolas", 10))
        self.editor.setStyleSheet(f"""
            QPlainTextEdit {{
                background-color: {AppPalette.BG_INPUT};
                color: {AppPalette.TEXT_PRIMARY};
                border: 1px solid {AppPalette.BORDER_COLOR};
                border-radius: 8px;
                padding: 12px;
                selection-background-color: {AppPalette.PRIMARY};
                selection-color: #FFFFFF;
            }}
        """)
        layout.addWidget(self.editor)

        # Footer Actions
        footer_row = QHBoxLayout()
        footer_row.setSpacing(10)

        reset_btn = QPushButton("Reset to Default")
        reset_btn.setFont(QFont("Segoe UI", 9))
        reset_btn.setStyleSheet(f"""
            QPushButton {{
                background-color: transparent;
                color: {AppPalette.TEXT_SECONDARY};
                border: 1px solid {AppPalette.BORDER_COLOR};
                border-radius: 6px;
                padding: 8px 14px;
            }}
            QPushButton:hover {{
                background-color: {AppPalette.BG_CARD_HOVER};
                color: {AppPalette.ERROR};
                border-color: {AppPalette.ERROR};
            }}
        """)
        reset_btn.clicked.connect(self._on_reset)
        footer_row.addWidget(reset_btn)

        footer_row.addStretch()

        copy_btn = QPushButton("Copy to Clipboard")
        copy_btn.setFont(QFont("Segoe UI", 9))
        copy_btn.setStyleSheet(f"""
            QPushButton {{
                background-color: {AppPalette.BG_CARD};
                color: {AppPalette.TEXT_PRIMARY};
                border: 1px solid {AppPalette.BORDER_COLOR};
                border-radius: 6px;
                padding: 8px 14px;
            }}
            QPushButton:hover {{
                background-color: {AppPalette.BG_CARD_HOVER};
                border-color: {AppPalette.BORDER_LIGHT};
            }}
        """)
        copy_btn.clicked.connect(self._on_copy)
        footer_row.addWidget(copy_btn)

        save_btn = QPushButton("Save Changes")
        save_btn.setFont(QFont("Segoe UI", 9, QFont.Bold))
        save_btn.setStyleSheet(f"""
            QPushButton {{
                background-color: {AppPalette.PRIMARY};
                color: #FFFFFF;
                border: none;
                border-radius: 6px;
                padding: 8px 18px;
            }}
            QPushButton:hover {{
                background-color: {AppPalette.PRIMARY_HOVER};
            }}
        """)
        save_btn.clicked.connect(self._on_save)
        footer_row.addWidget(save_btn)

        close_btn = QPushButton("Close")
        close_btn.setFont(QFont("Segoe UI", 9))
        close_btn.setStyleSheet(f"""
            QPushButton {{
                background-color: {AppPalette.BG_CARD};
                color: {AppPalette.TEXT_SECONDARY};
                border: 1px solid {AppPalette.BORDER_COLOR};
                border-radius: 6px;
                padding: 8px 14px;
            }}
            QPushButton:hover {{
                background-color: {AppPalette.BG_CARD_HOVER};
                color: {AppPalette.TEXT_PRIMARY};
            }}
        """)
        close_btn.clicked.connect(self.accept)
        footer_row.addWidget(close_btn)

        layout.addLayout(footer_row)

    def _update_badge_style(self):
        if self.is_custom:
            self.badge_label.setText("User Custom Override")
            self.badge_label.setStyleSheet(f"""
                background-color: {AppPalette.WARNING};
                color: #FFFFFF;
                border-radius: 10px;
                padding: 2px 8px;
            """)
        else:
            self.badge_label.setText("Default Bundled")
            self.badge_label.setStyleSheet(f"""
                background-color: {AppPalette.BG_SURFACE};
                color: {AppPalette.TEXT_SECONDARY};
                border: 1px solid {AppPalette.BORDER_COLOR};
                border-radius: 10px;
                padding: 2px 8px;
            """)

    def _on_copy(self):
        text = self.editor.toPlainText()
        QApplication.clipboard().setText(text)
        if self.on_status_change:
            self.on_status_change("System prompt copied to clipboard!", "info")

    def _on_save(self):
        content = self.editor.toPlainText()
        ok = self.pm.save_user_prompt(self.copilot_key, content)
        if ok:
            self.is_custom = True
            self._update_badge_style()
            if self.on_status_change:
                self.on_status_change("Saved custom system prompt override in AppData", "success")

    def _on_reset(self):
        self.pm.reset_to_default(self.copilot_key)
        default_text = self.pm.get_default_prompt(self.copilot_key)
        self.editor.setPlainText(default_text)
        self.is_custom = False
        self._update_badge_style()
        if self.on_status_change:
            self.on_status_change("Restored factory default prompt", "info")


def open_prompt_dialog(
    copilot_key: str,
    on_status_change: Optional[Callable[[str, str], None]] = None,
    parent: Optional[QWidget] = None,
):
    """Helper to instantiate and show PromptDialog."""
    dialog = PromptDialog(copilot_key, on_status_change, parent)
    dialog.exec()
