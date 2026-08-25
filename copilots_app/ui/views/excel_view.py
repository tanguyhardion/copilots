"""
Excel Copilot View in PySide6: Workbook Explorer, LLM Context Generator, JSON Action Protocol Executor, and Backup Manager.
"""

import os
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
    QFileDialog,
    QStackedWidget,
    QPlainTextEdit,
    QSplitter,
)
from PySide6.QtCore import Qt, Signal, QObject
from PySide6.QtGui import QFont

from copilots_app.core.theme import AppPalette
from copilots_app.ui.components import AppHeader, StatusBar, CodeEditor, ActionButton, MetricCard
from copilots_app.ui.prompt_dialog import open_prompt_dialog
from copilots_app.core.prompt_manager import PromptManager
from copilots_app.services.excel.analyzer.workbook_analyzer import WorkbookAnalyzer
from copilots_app.services.excel.executor.executor_main import ActionExecutor
from copilots_app.services.excel.protocol.action_parser import ActionParser
from copilots_app.services.excel.models.protocol import ActionProtocol, ExecutionStatus

SAMPLE_ACTION_JSON = """\
{
  "intent": "MODIFY_WORKBOOK",
  "version": "1.0",
  "actions": [
    {
      "action": "SET_CELL_VALUE",
      "sheet": "Summary",
      "cell": "B2",
      "value": "Updated by Copilot Suite"
    },
    {
      "action": "FORMAT_CELL",
      "sheet": "Summary",
      "range": "B2:D2",
      "bold": true,
      "fill_color": "E2EFDA",
      "font_color": "375623"
    }
  ]
}
"""


class WorkerSignals(QObject):
    status = Signal(str, str, bool)
    log = Signal(str, str)
    finished = Signal()


class ExcelView(QWidget):
    """Unified Excel Copilot interface in PySide6."""

    def __init__(self, parent: Optional[QWidget] = None):
        super().__init__(parent)
        self.executor = ActionExecutor()
        self.current_file_path: Optional[str] = None
        self.current_model = None

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
            "excel",
            on_status_change=lambda msg, lvl: self.status_bar.set_status(msg, level=lvl),
            parent=self,
        ))

        demo_btn = QPushButton("Open Demo Workbook")
        demo_btn.setFont(QFont("Segoe UI", 9))
        demo_btn.setStyleSheet(f"""
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
        demo_btn.clicked.connect(self._load_demo_workbook)

        self.header = AppHeader(
            title="Excel Copilot",
            subtitle="Air-gapped semantic analysis, LLM context generation, and deterministic JSON action execution",
            icon_path="icons/excel.png",
            badge_text="openpyxl + semantic engine",
            badge_color=AppPalette.BRAND_EXCEL,
            actions=[prompt_btn, demo_btn],
            parent=self,
        )
        layout.addWidget(self.header)

        # Top Overview Metric Cards
        metric_container = QWidget()
        m_layout = QHBoxLayout(metric_container)
        m_layout.setContentsMargins(16, 12, 16, 12)
        m_layout.setSpacing(10)

        self.metric_file = MetricCard(title="Active File", value="None", icon_text="📄", color=AppPalette.BRAND_EXCEL)
        self.metric_sheets = MetricCard(title="Sheets", value="0", icon_text="⊞", color=AppPalette.PRIMARY)
        self.metric_tables = MetricCard(title="Tables", value="0", icon_text="▦", color=AppPalette.INFO)
        self.metric_formulas = MetricCard(title="Formulas", value="0", icon_text="fx", color=AppPalette.WARNING)

        m_layout.addWidget(self.metric_file)
        m_layout.addWidget(self.metric_sheets)
        m_layout.addWidget(self.metric_tables)
        m_layout.addWidget(self.metric_formulas)
        layout.addWidget(metric_container)

        # Main Split Body
        split_widget = QWidget()
        split_layout = QHBoxLayout(split_widget)
        split_layout.setContentsMargins(16, 0, 16, 16)
        split_layout.setSpacing(12)

        # Left Panel (Explorer / Prompt Context tabs)
        left_panel = QFrame()
        left_panel.setStyleSheet(f"""
            background-color: {AppPalette.BG_CARD};
            border: 1px solid {AppPalette.BORDER_COLOR};
            border-radius: 8px;
        """)
        lp_layout = QVBoxLayout(left_panel)
        lp_layout.setContentsMargins(8, 8, 8, 8)
        lp_layout.setSpacing(8)

        # Tab toggle buttons
        tab_row = QHBoxLayout()
        tab_row.setSpacing(6)

        self.btn_tab_explorer = QPushButton("Workbook Explorer")
        self.btn_tab_explorer.setFont(QFont("Segoe UI", 9, QFont.Bold))
        self.btn_tab_prompt = QPushButton("LLM Prompt Context")
        self.btn_tab_prompt.setFont(QFont("Segoe UI", 9))

        tab_row.addWidget(self.btn_tab_explorer)
        tab_row.addWidget(self.btn_tab_prompt)
        lp_layout.addLayout(tab_row)

        self.left_stack = QStackedWidget()

        # Page 1: Explorer Scroll
        self.sheets_scroll = QScrollArea()
        self.sheets_scroll.setWidgetResizable(True)
        self.sheets_scroll.setStyleSheet("border: none; background: transparent;")
        self.sheets_container = QWidget()
        self.sheets_layout = QVBoxLayout(self.sheets_container)
        self.sheets_layout.setContentsMargins(4, 4, 4, 4)
        self.sheets_layout.setSpacing(6)
        self.sheets_layout.addStretch()
        self.sheets_scroll.setWidget(self.sheets_container)
        self.left_stack.addWidget(self.sheets_scroll)

        # Page 2: Prompt Context Area
        prompt_page = QWidget()
        pp_layout = QVBoxLayout(prompt_page)
        pp_layout.setContentsMargins(0, 0, 0, 0)
        pp_layout.setSpacing(6)

        pp_toolbar = QHBoxLayout()
        pp_title = QLabel("Semantic Model & System Prompt")
        pp_title.setFont(QFont("Segoe UI", 8, QFont.Bold))
        pp_title.setStyleSheet(f"color: {AppPalette.TEXT_MUTED};")
        pp_toolbar.addWidget(pp_title)
        pp_toolbar.addStretch()

        copy_prompt_btn = QPushButton("Copy All Context")
        copy_prompt_btn.setFont(QFont("Segoe UI", 8))
        copy_prompt_btn.setStyleSheet(f"""
            QPushButton {{
                background-color: {AppPalette.BG_SURFACE};
                color: {AppPalette.TEXT_PRIMARY};
                border: 1px solid {AppPalette.BORDER_COLOR};
                border-radius: 4px;
                padding: 3px 8px;
            }}
            QPushButton:hover {{
                background-color: {AppPalette.BG_CARD_HOVER};
            }}
        """)
        copy_prompt_btn.clicked.connect(self._copy_prompt_context)
        pp_toolbar.addWidget(copy_prompt_btn)
        pp_layout.addLayout(pp_toolbar)

        self.prompt_context_text = QPlainTextEdit()
        self.prompt_context_text.setReadOnly(True)
        self.prompt_context_text.setPlainText("Open a workbook to generate LLM context and schema model...")
        self.prompt_context_text.setFont(QFont("Consolas", 9))
        self.prompt_context_text.setStyleSheet(f"""
            QPlainTextEdit {{
                background-color: {AppPalette.BG_INPUT};
                color: {AppPalette.TEXT_PRIMARY};
                border: 1px solid {AppPalette.BORDER_COLOR};
                border-radius: 6px;
                padding: 8px;
            }}
        """)
        pp_layout.addWidget(self.prompt_context_text)
        self.left_stack.addWidget(prompt_page)

        lp_layout.addWidget(self.left_stack)

        self.btn_tab_explorer.clicked.connect(lambda: self._switch_left_tab(0))
        self.btn_tab_prompt.clicked.connect(lambda: self._switch_left_tab(1))
        self._update_tab_button_styles(0)

        split_layout.addWidget(left_panel, 4)

        # Right Panel (Protocol JSON Editor + Validation/Log Panel)
        right_panel = QWidget()
        rp_layout = QVBoxLayout(right_panel)
        rp_layout.setContentsMargins(0, 0, 0, 0)
        rp_layout.setSpacing(8)

        self.protocol_editor = CodeEditor(
            value=SAMPLE_ACTION_JSON,
            hint_text="Enter Action Protocol JSON here...",
            parent=self,
        )
        rp_layout.addWidget(self.protocol_editor, 1)

        # Validation log panel
        val_frame = QFrame()
        val_frame.setFixedHeight(120)
        val_frame.setStyleSheet(f"""
            background-color: {AppPalette.BG_CARD};
            border: 1px solid {AppPalette.BORDER_COLOR};
            border-radius: 8px;
        """)
        vf_layout = QVBoxLayout(val_frame)
        vf_layout.setContentsMargins(12, 8, 12, 8)
        vf_layout.setSpacing(4)

        vf_title = QLabel("Validation & Execution Log")
        vf_title.setFont(QFont("Segoe UI", 9, QFont.Bold))
        vf_title.setStyleSheet(f"color: {AppPalette.TEXT_PRIMARY};")
        vf_layout.addWidget(vf_title)

        val_scroll = QScrollArea()
        val_scroll.setWidgetResizable(True)
        val_scroll.setStyleSheet("border: none; background: transparent;")
        self.log_text = QLabel("No execution logs yet.")
        self.log_text.setFont(QFont("Consolas", 9))
        self.log_text.setStyleSheet(f"color: {AppPalette.TEXT_MUTED};")
        self.log_text.setWordWrap(True)
        val_scroll.setWidget(self.log_text)
        vf_layout.addWidget(val_scroll)

        rp_layout.addWidget(val_frame)
        split_layout.addWidget(right_panel, 6)

        layout.addWidget(split_widget, 1)

        # Bottom Action Bar
        action_bar = QFrame()
        action_bar.setStyleSheet(f"""
            background-color: {AppPalette.BG_SURFACE};
            border-top: 1px solid {AppPalette.BORDER_COLOR};
        """)
        act_layout = QHBoxLayout(action_bar)
        act_layout.setContentsMargins(16, 12, 16, 12)
        act_layout.setSpacing(10)

        self.open_file_btn = ActionButton(
            text="Open Workbook",
            color=AppPalette.PRIMARY,
        )
        self.open_file_btn.clicked.connect(self._on_open_file_dialog)
        act_layout.addWidget(self.open_file_btn)

        self.execute_btn = ActionButton(
            text="Execute Protocol",
            color=AppPalette.BRAND_EXCEL,
            tooltip="Validates and applies the JSON action protocol with automatic backup",
        )
        self.execute_btn.clicked.connect(self._on_execute_protocol)
        act_layout.addWidget(self.execute_btn)

        self.backup_btn = ActionButton(
            text="View Backups",
            color="#4A4F57",
            tooltip="View safe automated snapshot backups",
        )
        self.backup_btn.clicked.connect(self._on_view_backups)
        act_layout.addWidget(self.backup_btn)

        act_layout.addStretch()
        layout.addWidget(action_bar)

        # Status Bar
        self.status_bar = StatusBar(default_text="Ready — open an Excel workbook to begin", parent=self)
        layout.addWidget(self.status_bar)

    def _switch_left_tab(self, index: int):
        self.left_stack.setCurrentIndex(index)
        self._update_tab_button_styles(index)

    def _update_tab_button_styles(self, active_index: int):
        if active_index == 0:
            self.btn_tab_explorer.setStyleSheet(f"""
                QPushButton {{
                    background-color: {AppPalette.BG_CARD_HOVER};
                    color: {AppPalette.TEXT_PRIMARY};
                    border: 1px solid {AppPalette.BORDER_LIGHT};
                    border-radius: 5px;
                    padding: 5px 10px;
                }}
            """)
            self.btn_tab_prompt.setStyleSheet(f"""
                QPushButton {{
                    background-color: transparent;
                    color: {AppPalette.TEXT_MUTED};
                    border: none;
                    padding: 5px 10px;
                }}
                QPushButton:hover {{
                    color: {AppPalette.TEXT_PRIMARY};
                }}
            """)
        else:
            self.btn_tab_prompt.setStyleSheet(f"""
                QPushButton {{
                    background-color: {AppPalette.BG_CARD_HOVER};
                    color: {AppPalette.TEXT_PRIMARY};
                    border: 1px solid {AppPalette.BORDER_LIGHT};
                    border-radius: 5px;
                    padding: 5px 10px;
                }}
            """)
            self.btn_tab_explorer.setStyleSheet(f"""
                QPushButton {{
                    background-color: transparent;
                    color: {AppPalette.TEXT_MUTED};
                    border: none;
                    padding: 5px 10px;
                }}
                QPushButton:hover {{
                    color: {AppPalette.TEXT_PRIMARY};
                }}
            """)

    def _load_demo_workbook(self):
        demo_path = os.path.abspath("excel-copilot/Financial_Dashboard.xlsx")
        if not os.path.exists(demo_path):
            demo_path = os.path.abspath("Financial_Dashboard.xlsx")
        if os.path.exists(demo_path):
            self.load_workbook(demo_path)
        else:
            self.status_bar.set_status(f"Demo file not found at {demo_path}", level="warning")

    def _on_open_file_dialog(self):
        file_path, _ = QFileDialog.getOpenFileName(
            self, "Open Excel Workbook", "", "Excel Files (*.xlsx *.xlsm *.xltx *.xltm)"
        )
        if file_path:
            self.load_workbook(file_path)

    def load_workbook(self, file_path: str):
        try:
            self.current_file_path = os.path.abspath(file_path)
            self.status_bar.set_status("Analyzing workbook structure…", level="info", loading=True)
            self.current_model = WorkbookAnalyzer.analyze(self.current_file_path)

            # Update Metrics
            self.metric_file.set_value(os.path.basename(self.current_file_path))
            self.metric_sheets.set_value(str(len(self.current_model.worksheets)))

            total_tables = sum(len(s.tables) for s in self.current_model.worksheets)
            self.metric_tables.set_value(str(total_tables))

            total_formulas = sum(s.formulas_count for s in self.current_model.worksheets)
            self.metric_formulas.set_value(str(total_formulas))

            # Clear sheets container
            while self.sheets_layout.count() > 1:
                child = self.sheets_layout.takeAt(0)
                if child.widget():
                    child.widget().deleteLater()

            # Populate sheets list
            for sheet in self.current_model.worksheets:
                card = QFrame()
                card.setStyleSheet(f"""
                    background-color: {AppPalette.BG_CARD_HOVER};
                    border: 1px solid {AppPalette.BORDER_COLOR};
                    border-radius: 6px;
                """)
                c_layout = QVBoxLayout(card)
                c_layout.setContentsMargins(10, 8, 10, 8)
                c_layout.setSpacing(3)

                header_row = QHBoxLayout()
                h_name = QLabel(f"⊞ {sheet.name}")
                h_name.setFont(QFont("Segoe UI", 9, QFont.Bold))
                h_name.setStyleSheet(f"color: {AppPalette.TEXT_PRIMARY};")
                header_row.addWidget(h_name)

                dim_badge = QLabel(f"{sheet.max_row}x{sheet.max_column}")
                dim_badge.setFont(QFont("Segoe UI", 8))
                dim_badge.setStyleSheet(f"""
                    background-color: {AppPalette.BG_INPUT};
                    color: {AppPalette.TEXT_MUTED};
                    border-radius: 4px;
                    padding: 1px 6px;
                """)
                header_row.addWidget(dim_badge)
                header_row.addStretch()
                c_layout.addLayout(header_row)

                info_label = QLabel(f"Tables: {len(sheet.tables)}  |  Charts: {len(sheet.charts)}  |  Formulas: {sheet.formulas_count}")
                info_label.setFont(QFont("Segoe UI", 8))
                info_label.setStyleSheet(f"color: {AppPalette.TEXT_SECONDARY};")
                c_layout.addWidget(info_label)

                self.sheets_layout.insertWidget(self.sheets_layout.count() - 1, card)

            # Generate LLM Context
            context_str = f"# EXCEL WORKBOOK SEMANTIC CONTEXT\nFile: {os.path.basename(self.current_file_path)}\n\n"
            for s in self.current_model.worksheets:
                context_str += f"## Sheet: {s.name} (Rows: {s.max_row}, Cols: {s.max_column})\n"
                if s.tables:
                    context_str += f"- Tables: {', '.join(t.name for t in s.tables)}\n"
                if s.formulas_count:
                    context_str += f"- Formulas Count: {s.formulas_count}\n"
            self.prompt_context_text.setPlainText(context_str)

            self.status_bar.set_status(f"✓ Loaded '{os.path.basename(self.current_file_path)}' successfully", level="success")
        except Exception as err:
            self.status_bar.set_status(f"Failed to load workbook: {err}", level="error")

    def _copy_prompt_context(self):
        system_prompt = PromptManager().get_prompt("excel")
        full_text = f"{system_prompt}\n\n================================================================================\nCURRENT WORKBOOK CONTEXT\n================================================================================\n{self.prompt_context_text.toPlainText()}"
        from PySide6.QtWidgets import QApplication
        QApplication.clipboard().setText(full_text)
        self.status_bar.set_status("Copied System Prompt + Workbook Context to clipboard!", level="success")

    def _on_execute_protocol(self):
        if not self.current_file_path:
            self.status_bar.set_status("Please open or load a workbook first", level="warning")
            return

        json_text = self.protocol_editor.get_value().strip()
        if not json_text:
            self.status_bar.set_status("Action Protocol JSON is empty", level="warning")
            return

        self.status_bar.set_status("Validating and executing Action Protocol…", level="info", loading=True)

        signals = WorkerSignals()
        signals.status.connect(self.status_bar.set_status)
        signals.log.connect(self._update_log)

        def worker():
            try:
                protocol = ActionParser.parse(json_text)
                res, updated_model = self.executor.execute(
                    protocol=protocol,
                    file_path=self.current_file_path,
                    model=self.current_model,
                )

                log_lines = []
                if res.status == ExecutionStatus.SUCCESS:
                    log_lines.append(f"✓ SUCCESS: Executed {res.actions_executed} actions.")
                    if res.objects_modified:
                        log_lines.append(f"Modified: {', '.join(res.objects_modified)}")
                    signals.status.emit(f"✓ Protocol executed successfully! ({res.actions_executed} actions)", "success", False)
                else:
                    log_lines.append(f"❌ {res.status.value}")
                    for err in res.errors:
                        log_lines.append(f"• Error: {err}")
                    signals.status.emit(f"Execution finished with status: {res.status.value}", "warning", False)

                for w in res.warnings:
                    log_lines.append(f"⚠ Warning: {w}")

                log_content = "\n".join(log_lines)
                color = AppPalette.SUCCESS if res.status == ExecutionStatus.SUCCESS else AppPalette.WARNING
                signals.log.emit(log_content, color)
            except Exception as err:
                signals.log.emit(f"Parse/Execution Error:\n{str(err)}", AppPalette.ERROR)
                signals.status.emit(f"Execution error: {err}", "error", False)

        threading.Thread(target=worker, daemon=True).start()

    def _update_log(self, text: str, color: str):
        self.log_text.setText(text)
        self.log_text.setStyleSheet(f"color: {color};")

    def _on_view_backups(self):
        backups = self.executor.backup_manager.list_backups()
        dialog = QDialog(self)
        dialog.setWindowTitle("Automated Backups")
        dialog.resize(520, 320)
        dialog.setStyleSheet(f"background-color: {AppPalette.BG_DARK};")

        dlg_layout = QVBoxLayout(dialog)
        dlg_layout.setContentsMargins(20, 20, 20, 20)
        dlg_layout.setSpacing(12)

        title = QLabel("Available Automatic Backups:")
        title.setFont(QFont("Segoe UI", 11, QFont.Bold))
        title.setStyleSheet(f"color: {AppPalette.TEXT_PRIMARY};")
        dlg_layout.addWidget(title)

        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setStyleSheet("border: none; background: transparent;")

        c_widget = QWidget()
        c_layout = QVBoxLayout(c_widget)
        c_layout.setSpacing(6)

        if not backups:
            no_b = QLabel("No backup snapshots found yet.")
            no_b.setStyleSheet(f"color: {AppPalette.TEXT_MUTED};")
            c_layout.addWidget(no_b)
        else:
            for b in backups:
                b_label = QLabel(f"• {b['timestamp']} — {os.path.basename(b.get('source_path', ''))} ({b['size_bytes']} bytes)")
                b_label.setFont(QFont("Consolas", 9))
                b_label.setStyleSheet(f"color: {AppPalette.TEXT_SECONDARY};")
                c_layout.addWidget(b_label)

        c_layout.addStretch()
        scroll.setWidget(c_widget)
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
