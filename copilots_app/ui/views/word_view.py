"""
Word Copilot View in PySide6: Rich DSL & Edit-Mode Editor, Build & Open, Insert, Apply Edits, Extract.
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
from copilots_app.services.word.connector import WordConnector
from copilots_app.services.word.editor import WordEditor
from copilots_app.services.word.extractor import WordExtractor
from copilots_app.services.word.colors import refresh_word_theme_colors
from copilots_app.services.word.dsl.parser import parse_dsl_pages
from copilots_app.services.word.dsl.edit_parser import parse_edit_dsl
from copilots_app.services.word.gui.sample import SAMPLE_DSL

WORD_SAMPLES = {
    "Complete Demo": SAMPLE_DSL,
    "Executive Report": """\
// Document setup
page size=a4 orientation=portrait margin=54,54,54,54

h1 align=center color=a4 | "Quarterly Strategy Review"
h3 align=center color=a2 | "Digital Operations & Modernization Program"

hr color=a3 weight=2

p spacing_after=8 | "This executive summary details the operational performance, technical deliverables, and roadmap milestones accomplished during the current cycle."

// Deliverables Table
table width=100% header_fill=a4 header_text_color=#FFFFFF text_color=t1 border_color=a3 border_weight=1 header_bold=true
cols=30%,45%,25%
header="Stream","Strategic Objective","Status"
row="Cloud Architecture","Scalable microservices migration","✓ Complete"
row="Automation Suite","Desktop Copilots unification in PySide6","✓ Production"
row="Data Quality","Automated compliance validation","✓ Verified"

br

h2 color=a4 | "Key Recommendations"
ul indent=1
  item | "Maintain strict deterministic data validation across all output formats."
  item | "Accelerate user enablement sessions with interactive cheatsheets."
  item | "Continuous integration testing for COM office automation hooks."
""",
    "Edit Mode Example": """\
// Target active document in Microsoft Word
edit target=active

// Search and replace or update headings
replace find="Draft" replace="Final Approved Version"
insert_before find="Key Recommendations" | "Note: All milestones reviewed by Executive Sponsor."
""",
}


class WorkerSignals(QObject):
    status = Signal(str, str, bool)
    finished = Signal()


class WordView(QWidget):
    """Unified Word Copilot interface in PySide6."""

    def __init__(self, parent: Optional[QWidget] = None):
        super().__init__(parent)
        self.connector = WordConnector()
        self.extractor = WordExtractor()
        self.editor = WordEditor()

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
            "word",
            on_status_change=lambda msg, lvl: self.status_bar.set_status(msg, level=lvl),
            parent=self,
        ))

        syntax_btn = QPushButton("Word DSL Syntax")
        syntax_btn.setFont(QFont("Segoe UI", 9))
        syntax_btn.setStyleSheet(f"""
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
        syntax_btn.clicked.connect(self._show_syntax_help)

        self.header = AppHeader(
            title="Word Copilot",
            subtitle="Build formatted documents, insert rich content at cursor, or live-edit active Word documents",
            icon_path="icons/word.png",
            badge_text="Word COM + python-docx",
            badge_color=AppPalette.BRAND_WORD,
            actions=[prompt_btn, syntax_btn],
            parent=self,
        )
        layout.addWidget(self.header)

        # Central Code Editor Area
        editor_container = QWidget()
        e_layout = QVBoxLayout(editor_container)
        e_layout.setContentsMargins(16, 16, 16, 16)
        e_layout.setSpacing(0)

        self.code_editor = CodeEditor(
            value=SAMPLE_DSL,
            hint_text="Paste Word build DSL (h1, p, table, ul, icon...) or edit DSL (edit target=active...)",
            samples=WORD_SAMPLES,
            on_sample_selected=self._on_sample_loaded,
            parent=self,
        )
        e_layout.addWidget(self.code_editor)
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

        self.open_btn = ActionButton(
            text="Build & Open in Word",
            color=AppPalette.BRAND_WORD,
            tooltip="Generates document from DSL and opens it in Microsoft Word",
        )
        self.open_btn.clicked.connect(self._on_build_and_open)
        act_layout.addWidget(self.open_btn)

        self.insert_btn = ActionButton(
            text="Insert at Cursor",
            color=AppPalette.SUCCESS,
            tooltip="Injects parsed elements directly at the current selection cursor in Word",
        )
        self.insert_btn.clicked.connect(self._on_insert_at_cursor)
        act_layout.addWidget(self.insert_btn)

        self.apply_btn = ActionButton(
            text="Apply Edits to Word",
            color="#5C3D8B",
            tooltip="Applies targeted edit operations to the active Word document",
        )
        self.apply_btn.clicked.connect(self._on_apply_edits)
        act_layout.addWidget(self.apply_btn)

        self.extract_btn = ActionButton(
            text="Extract DSL from Word",
            color=AppPalette.WARNING,
            tooltip="Reads the active Microsoft Word document and converts it to DSL in the editor",
        )
        self.extract_btn.clicked.connect(self._on_extract_dsl)
        act_layout.addWidget(self.extract_btn)

        self.save_btn = ActionButton(
            text="Save .docx",
            color="#4A4F57",
            tooltip="Saves DSL directly to a .docx file",
        )
        self.save_btn.clicked.connect(self._on_save_docx)
        act_layout.addWidget(self.save_btn)

        act_layout.addStretch()
        layout.addWidget(action_bar)

        # Bottom Status Bar
        self.status_bar = StatusBar(default_text="Ready — write build DSL or edit-mode DSL", parent=self)
        layout.addWidget(self.status_bar)

    def _set_buttons_state(self, disabled: bool):
        for btn in (self.open_btn, self.insert_btn, self.apply_btn, self.extract_btn, self.save_btn):
            btn.setDisabled(disabled)

    def _on_sample_loaded(self, sample_name: str):
        self.status_bar.set_status(f"Loaded Word template: '{sample_name}'", level="info")

    def _get_pages(self):
        text = self.code_editor.get_value().strip()
        if not text:
            return [[]]
        try:
            refresh_word_theme_colors()
        except Exception:
            pass
        return parse_dsl_pages(text)

    def _get_edit_ops(self):
        text = self.code_editor.get_value().strip()
        try:
            refresh_word_theme_colors()
        except Exception:
            pass
        return parse_edit_dsl(text)

    def _on_build_and_open(self):
        pages = self._get_pages()
        if not any(pages):
            self.status_bar.set_status("DSL is empty — nothing to build", level="warning")
            return

        self._set_buttons_state(True)
        self.status_bar.set_status("Building document and opening Word…", level="info", loading=True)

        signals = WorkerSignals()
        signals.status.connect(self.status_bar.set_status)
        signals.finished.connect(lambda: self._set_buttons_state(False))

        def worker():
            try:
                self.connector.build_and_open(
                    pages,
                    status_cb=lambda m: signals.status.emit(m, "info", True),
                )
                signals.status.emit("✓ Document built and opened in Microsoft Word!", "success", False)
            except Exception as err:
                signals.status.emit(f"Build failed: {err}", "error", False)
            finally:
                signals.finished.emit()

        threading.Thread(target=worker, daemon=True).start()

    def _on_insert_at_cursor(self):
        pages = self._get_pages()
        if not any(pages):
            self.status_bar.set_status("DSL is empty — nothing to insert", level="warning")
            return

        self._set_buttons_state(True)
        self.status_bar.set_status("Inserting content at Word cursor…", level="info", loading=True)

        signals = WorkerSignals()
        signals.status.connect(self.status_bar.set_status)
        signals.finished.connect(lambda: self._set_buttons_state(False))

        def worker():
            try:
                self.connector.insert_at_cursor(
                    pages,
                    status_cb=lambda m: signals.status.emit(m, "info", True),
                )
                signals.status.emit("✓ Content successfully inserted at cursor in Word!", "success", False)
            except Exception as err:
                signals.status.emit(f"Insert failed: {err}", "error", False)
            finally:
                signals.finished.emit()

        threading.Thread(target=worker, daemon=True).start()

    def _on_apply_edits(self):
        ops = self._get_edit_ops()
        if not ops:
            self.status_bar.set_status("No valid edit operations found in DSL", level="warning")
            return

        self._set_buttons_state(True)
        self.status_bar.set_status("Applying edits to active Word document…", level="info", loading=True)

        signals = WorkerSignals()
        signals.status.connect(self.status_bar.set_status)
        signals.finished.connect(lambda: self._set_buttons_state(False))

        def worker():
            try:
                self.editor.apply_edits(
                    ops,
                    status_cb=lambda m: signals.status.emit(m, "info", True),
                )
                signals.status.emit("✓ Edit operations applied successfully!", "success", False)
            except Exception as err:
                signals.status.emit(f"Apply edits failed: {err}", "error", False)
            finally:
                signals.finished.emit()

        threading.Thread(target=worker, daemon=True).start()

    def _on_extract_dsl(self):
        self._set_buttons_state(True)
        self.status_bar.set_status("Extracting content from active Word document…", level="info", loading=True)

        signals = WorkerSignals()
        signals.status.connect(self.status_bar.set_status)
        signals.finished.connect(lambda: self._set_buttons_state(False))

        def worker():
            try:
                dsl_text = self.extractor.extract(
                    status_cb=lambda m: signals.status.emit(m, "info", True),
                )
                if dsl_text:
                    self.code_editor.set_value(dsl_text)
                    signals.status.emit("✓ Successfully extracted DSL from Word!", "success", False)
                else:
                    signals.status.emit("No content extracted from active document", "warning", False)
            except Exception as err:
                signals.status.emit(f"Extraction failed: {err}", "error", False)
            finally:
                signals.finished.emit()

        threading.Thread(target=worker, daemon=True).start()

    def _on_save_docx(self):
        pages = self._get_pages()
        if not any(pages):
            self.status_bar.set_status("DSL is empty — nothing to save", level="warning")
            return

        self._set_buttons_state(True)
        self.status_bar.set_status("Saving document as .docx…", level="info", loading=True)

        signals = WorkerSignals()
        signals.status.connect(self.status_bar.set_status)
        signals.finished.connect(lambda: self._set_buttons_state(False))

        def worker():
            try:
                out_file = "Generated_Word_Document.docx"
                self.connector.save_to_file(
                    pages,
                    out_file,
                    status_cb=lambda m: signals.status.emit(m, "info", True),
                )
                signals.status.emit(f"✓ Saved document to '{out_file}'", "success", False)
            except Exception as err:
                signals.status.emit(f"Save failed: {err}", "error", False)
            finally:
                signals.finished.emit()

        threading.Thread(target=worker, daemon=True).start()

    def _show_syntax_help(self):
        dialog = QDialog(self)
        dialog.setWindowTitle("Word DSL Specification")
        dialog.resize(600, 420)
        dialog.setStyleSheet(f"background-color: {AppPalette.BG_DARK};")

        dlg_layout = QVBoxLayout(dialog)
        dlg_layout.setContentsMargins(20, 20, 20, 20)
        dlg_layout.setSpacing(12)

        title = QLabel("Word DSL Specification")
        title.setFont(QFont("Segoe UI", 12, QFont.Bold))
        title.setStyleSheet(f"color: {AppPalette.TEXT_PRIMARY};")
        dlg_layout.addWidget(title)

        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setStyleSheet("border: none;")

        content_widget = QWidget()
        c_layout = QVBoxLayout(content_widget)
        c_layout.setSpacing(8)

        syntax_text = """
<b>Structure Directives:</b><br>
<code>page size=letter|a4 orientation=portrait|landscape margin=72,72,72,72</code><br>
<code>h1..h6 align=left|center|right color=a1..a6 | "Heading Text"</code><br>
<code>p spacing_after=6 | "Text with " + "bold" bold=true + " parts"</code><br><br>
<b>Lists & Tables:</b><br>
<code>ul / ol indent=1 followed by indented item | "bullet"</code><br>
<code>table width=100% header_fill=a4 text_color=t1 ... cols=... header=... row=...</code><br><br>
<b>Edit DSL (for active documents):</b><br>
<code>edit target=active</code><br>
<code>replace find="Old" replace="New"</code><br>
<code>insert_before find="Target" | "New content"</code>
"""
        label = QLabel(syntax_text)
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
