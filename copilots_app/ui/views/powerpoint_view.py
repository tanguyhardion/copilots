"""
PowerPoint Copilot View in PySide6: DSL Editor, Quick Templates, Presentation Controls, and Cheatsheet.
"""

import threading
from typing import Optional
from PySide6.QtWidgets import (
    QWidget,
    QVBoxLayout,
    QHBoxLayout,
    QFrame,
    QPushButton,
    QDialog,
    QLabel,
    QScrollArea,
)
from PySide6.QtCore import Qt, Signal, QObject
from PySide6.QtGui import QFont

from copilots_app.core.theme import AppPalette
from copilots_app.ui.components import AppHeader, StatusBar, CodeEditor, ActionButton
from copilots_app.ui.prompt_dialog import open_prompt_dialog
from copilots_app.services.powerpoint import (
    PPT_SAMPLES,
    PowerPointConnector,
    parse_dsl_slides,
    refresh_dsl_theme_colors,
)


class WorkerSignals(QObject):
    status = Signal(str, str, bool)
    finished = Signal()


class PowerPointView(QWidget):
    """Unified PowerPoint Copilot interface in PySide6."""

    def __init__(self, parent: Optional[QWidget] = None):
        super().__init__(parent)
        self.connector = PowerPointConnector()

        layout = QVBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(0)

        # Header Action Buttons
        prompt_btn = QPushButton("System Prompt")
        prompt_btn.setFont(QFont("Segoe UI", 9))
        prompt_btn.setStyleSheet(f"""
            QPushButton {{
                background-color: {AppPalette.BG_CARD};
                color: {AppPalette.TEXT_SECONDARY};
                border: 1px solid {AppPalette.BORDER_COLOR};
                border-radius: 6px;
                padding: 6px 12px;
            }}
            QPushButton:hover {{
                background-color: {AppPalette.BG_CARD_HOVER};
                color: {AppPalette.TEXT_PRIMARY};
            }}
        """)
        prompt_btn.clicked.connect(lambda: open_prompt_dialog(
            "powerpoint",
            on_status_change=lambda msg, lvl: self.status_bar.set_status(msg, level=lvl),
            parent=self,
        ))

        cheatsheet_btn = QPushButton("DSL Cheatsheet")
        cheatsheet_btn.setFont(QFont("Segoe UI", 9))
        cheatsheet_btn.setStyleSheet(f"""
            QPushButton {{
                background-color: {AppPalette.BG_CARD};
                color: {AppPalette.TEXT_SECONDARY};
                border: 1px solid {AppPalette.BORDER_COLOR};
                border-radius: 6px;
                padding: 6px 12px;
            }}
            QPushButton:hover {{
                background-color: {AppPalette.BG_CARD_HOVER};
                color: {AppPalette.TEXT_PRIMARY};
            }}
        """)
        cheatsheet_btn.clicked.connect(self._show_cheatsheet)

        self.header = AppHeader(
            title="PowerPoint Copilot",
            subtitle="Generate slides, shapes, rich text, tables, and icons from clean DSL",
            icon_path="icons/powerpoint.png",
            badge_text="PowerPoint COM",
            badge_color=AppPalette.BRAND_PPT,
            actions=[prompt_btn, cheatsheet_btn],
            parent=self,
        )
        layout.addWidget(self.header)

        # Central DSL Editor Area
        editor_container = QWidget()
        e_layout = QVBoxLayout(editor_container)
        e_layout.setContentsMargins(16, 16, 16, 16)
        e_layout.setSpacing(0)

        self.editor = CodeEditor(
            value=PPT_SAMPLES.get("Overview Demo", ""),
            hint_text="Enter PowerPoint DSL here (rect, rounded_rect, table, icon, chevron...)",
            samples=PPT_SAMPLES,
            on_sample_selected=self._on_sample_loaded,
            parent=self,
        )
        e_layout.addWidget(self.editor)
        layout.addWidget(editor_container, 1)

        # Bottom Action Bar
        action_bar = QFrame()
        action_bar.setStyleSheet(f"""
            background-color: {AppPalette.BG_SURFACE};
            border-top: 1px solid {AppPalette.BORDER_COLOR};
        """)
        act_layout = QHBoxLayout(action_bar)
        act_layout.setContentsMargins(16, 12, 16, 12)
        act_layout.setSpacing(10)

        self.copy_btn = ActionButton(
            text="Copy to Clipboard",
            color=AppPalette.BRAND_PPT,
            tooltip="Creates shapes in background and copies them to clipboard for Ctrl+V in PowerPoint",
        )
        self.copy_btn.clicked.connect(self._on_copy_clipboard)
        act_layout.addWidget(self.copy_btn)

        self.insert_btn = ActionButton(
            text="Insert on Current Slide",
            color=AppPalette.SUCCESS,
            tooltip="Directly injects shapes onto the currently selected slide in PowerPoint",
        )
        self.insert_btn.clicked.connect(self._on_insert_current_slide)
        act_layout.addWidget(self.insert_btn)

        self.slide_btn = ActionButton(
            text="Create Full Slide(s)",
            color=AppPalette.PRIMARY,
            tooltip="Appends one or more new slides in the active PowerPoint presentation",
        )
        self.slide_btn.clicked.connect(self._on_create_full_slide)
        act_layout.addWidget(self.slide_btn)

        act_layout.addStretch()
        layout.addWidget(action_bar)

        # Bottom Status Bar
        self.status_bar = StatusBar(default_text="Ready — write or paste DSL and choose an action", parent=self)
        layout.addWidget(self.status_bar)

    def _set_buttons_state(self, disabled: bool):
        self.copy_btn.setDisabled(disabled)
        self.insert_btn.setDisabled(disabled)
        self.slide_btn.setDisabled(disabled)

    def _on_sample_loaded(self, sample_name: str):
        self.status_bar.set_status(f"Loaded sample template: '{sample_name}'", level="info")

    def _get_slides(self):
        text = self.editor.get_value().strip()
        if not text:
            return [[]]
        try:
            refresh_dsl_theme_colors()
        except Exception:
            pass
        return parse_dsl_slides(text)

    def _on_copy_clipboard(self):
        slides = self._get_slides()
        shapes = slides[0] if slides else []
        if not shapes:
            self.status_bar.set_status("DSL is empty — nothing to copy", level="warning")
            return

        self._set_buttons_state(True)
        self.status_bar.set_status("Building shapes for clipboard…", level="info", loading=True)

        signals = WorkerSignals()
        signals.status.connect(self.status_bar.set_status)
        signals.finished.connect(lambda: self._set_buttons_state(False))

        def worker():
            try:
                self.connector.create_shapes_and_copy(
                    shapes,
                    status_cb=lambda m: signals.status.emit(m, "info", True),
                )
                signals.status.emit("✓ Shapes copied to clipboard — switch to PowerPoint and press Ctrl+V", "success", False)
            except Exception as err:
                signals.status.emit(f"Clipboard copy failed: {err}", "error", False)
            finally:
                signals.finished.emit()

        threading.Thread(target=worker, daemon=True).start()

    def _on_insert_current_slide(self):
        slides = self._get_slides()
        shapes = slides[0] if slides else []
        if not shapes:
            self.status_bar.set_status("DSL is empty — nothing to insert", level="warning")
            return

        self._set_buttons_state(True)
        self.status_bar.set_status("Inserting shapes onto active slide…", level="info", loading=True)

        signals = WorkerSignals()
        signals.status.connect(self.status_bar.set_status)
        signals.finished.connect(lambda: self._set_buttons_state(False))

        def worker():
            try:
                self.connector.create_on_current_slide(
                    shapes,
                    status_cb=lambda m: signals.status.emit(m, "info", True),
                )
                signals.status.emit("✓ Shapes successfully inserted on active slide!", "success", False)
            except Exception as err:
                signals.status.emit(f"Insert failed: {err}", "error", False)
            finally:
                signals.finished.emit()

        threading.Thread(target=worker, daemon=True).start()

    def _on_create_full_slide(self):
        slides = self._get_slides()
        total_shapes = sum(len(s) for s in slides)
        if total_shapes == 0:
            self.status_bar.set_status("DSL is empty — nothing to build", level="warning")
            return

        self._set_buttons_state(True)
        self.status_bar.set_status("Creating new slide(s)…", level="info", loading=True)

        signals = WorkerSignals()
        signals.status.connect(self.status_bar.set_status)
        signals.finished.connect(lambda: self._set_buttons_state(False))

        def worker():
            try:
                self.connector.create_on_new_slide(
                    slides,
                    status_cb=lambda m: signals.status.emit(m, "info", True),
                )
                signals.status.emit(f"✓ Created {len(slides)} slide(s) ({total_shapes} shapes total) successfully!", "success", False)
            except Exception as err:
                signals.status.emit(f"Slide creation failed: {err}", "error", False)
            finally:
                signals.finished.emit()

        threading.Thread(target=worker, daemon=True).start()

    def _show_cheatsheet(self):
        dialog = QDialog(self)
        dialog.setWindowTitle("PowerPoint DSL Cheatsheet")
        dialog.resize(600, 420)
        dialog.setStyleSheet(f"background-color: {AppPalette.BG_DARK};")

        dlg_layout = QVBoxLayout(dialog)
        dlg_layout.setContentsMargins(20, 20, 20, 20)
        dlg_layout.setSpacing(12)

        title = QLabel("PowerPoint DSL Cheatsheet")
        title.setFont(QFont("Segoe UI", 12, QFont.Bold))
        title.setStyleSheet(f"color: {AppPalette.TEXT_PRIMARY};")
        dlg_layout.addWidget(title)

        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setStyleSheet("border: none;")

        content_widget = QWidget()
        c_layout = QVBoxLayout(content_widget)
        c_layout.setSpacing(8)

        cheatsheet_text = """
<b>Coordinate System: 960 × 540pt (16:9 Standard Slide)</b><br>
• Margins: left=48, right=912, top=48, bottom=492<br><br>
<b>Syntax Examples:</b><br>
<code>rect left=40 top=40 width=200 height=100 color=a1 | "Card Title" size=14 bold=true</code><br>
<code>rounded_rect left=260 top=40 width=200 height=100 border_radius=10 color=a6 outline=a1,2</code><br>
<code>icon name=chart-line style=solid left=60 top=60 width=32 height=32 color=a1</code><br>
<code>table left=48 top=200 width=864 height=200 cols=200,400,264 header="A","B","C" row="1","2","3"</code><br>
<code>line x1=48 y1=450 x2=912 y2=450 color=a3 weight=1.5 dash=solid</code><br><br>
<b>Slide Separator:</b> Use <code>---</code> on a separate line for multi-slide generation.
"""
        label = QLabel(cheatsheet_text)
        label.setWordWrap(True)
        label.setFont(QFont("Segoe UI", 10))
        label.setStyleSheet(f"color: {AppPalette.TEXT_SECONDARY}; line-height: 140%;")
        c_layout.addWidget(label)
        c_layout.addStretch()

        scroll.setWidget(content_widget)
        dlg_layout.addWidget(scroll)

        close_btn = QPushButton("Close")
        close_btn.setFont(QFont("Segoe UI", 9))
        close_btn.setStyleSheet(f"""
            QPushButton {{
                background-color: {AppPalette.BG_CARD};
                color: {AppPalette.TEXT_PRIMARY};
                border: 1px solid {AppPalette.BORDER_COLOR};
                border-radius: 6px;
                padding: 8px 16px;
            }}
            QPushButton:hover {{
                background-color: {AppPalette.BG_CARD_HOVER};
            }}
        """)
        close_btn.clicked.connect(dialog.accept)
        dlg_layout.addWidget(close_btn, 0, Qt.AlignRight)

        dialog.exec()
