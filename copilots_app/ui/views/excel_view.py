"""
Excel Copilot View: Workbook Explorer, LLM Context Generator, JSON Action Protocol Executor, and Backup Manager.
"""

import os
import json
import threading
import flet as ft
from typing import Optional, Dict, Any, List

from copilots_app.core.theme import AppPalette
from copilots_app.ui.components import AppHeader, StatusBar, CodeEditor, ActionButton, MetricCard
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


class ExcelView(ft.Container):
    """Unified Excel Copilot interface in Flet."""

    def __init__(self, page: ft.Page):
        self.app_page = page
        self.executor = ActionExecutor()
        self.current_file_path: Optional[str] = None
        self.current_model = None
        self.current_protocol: Optional[ActionProtocol] = None

        self.header = AppHeader(
            title="Excel Copilot",
            subtitle="Air-gapped semantic analysis, LLM context generation, and deterministic JSON action execution",
            icon_path="icons/excel.png",
            badge_text="openpyxl + semantic engine",
            badge_color=AppPalette.BRAND_EXCEL,
            actions=[
                ft.TextButton(
                    "Open Demo Workbook",
                    icon=ft.Icons.DESCRIPTION_OUTLINED,
                    on_click=self._load_demo_workbook,
                ),
            ],
        )

        self.status_bar = StatusBar(default_text="Ready — open an Excel workbook to begin")

        # Workbook Overview Cards
        self.metric_file = MetricCard(title="Active File", value="None", icon_name=ft.Icons.INSERT_DRIVE_FILE_OUTLINED, color=AppPalette.BRAND_EXCEL)
        self.metric_sheets = MetricCard(title="Sheets", value="0", icon_name=ft.Icons.GRID_ON, color=AppPalette.PRIMARY)
        self.metric_tables = MetricCard(title="Tables", value="0", icon_name=ft.Icons.TABLE_CHART_OUTLINED, color=AppPalette.INFO)
        self.metric_formulas = MetricCard(title="Formulas", value="0", icon_name=ft.Icons.FUNCTIONS, color=AppPalette.WARNING)

        overview_row = ft.Container(
            content=ft.Row(
                controls=[
                    self.metric_file,
                    self.metric_sheets,
                    self.metric_tables,
                    self.metric_formulas,
                ],
                spacing=10,
            ),
            padding=ft.Padding.symmetric(horizontal=16, vertical=12),
        )

        # Tabbed Workspace
        # Left Panel: Explorer / Context
        self.sheets_list = ft.ListView(spacing=6, expand=True, padding=8)
        self.prompt_context_text = ft.TextField(
            multiline=True,
            read_only=True,
            value="Open a workbook to generate LLM context and schema model...",
            text_size=12,
            text_style=ft.TextStyle(font_family="Consolas, monospace"),
            border=ft.InputBorder.NONE,
            filled=True,
            fill_color=AppPalette.BG_INPUT,
            expand=True,
        )

        left_tab_view = ft.Tabs(
            selected_index=0,
            animation_duration=200,
            tabs=[
                ft.Tab(
                    text="Workbook Explorer",
                    icon=ft.Icons.ACCOUNT_TREE_OUTLINED,
                    content=ft.Container(
                        content=self.sheets_list,
                        bgcolor=AppPalette.BG_CARD,
                        border=ft.Border.all(1, AppPalette.BORDER_COLOR),
                        border_radius=8,
                        padding=4,
                    ),
                ),
                ft.Tab(
                    text="LLM Prompt Context",
                    icon=ft.Icons.PSYCHOLOGY_OUTLINED,
                    content=ft.Container(
                        content=ft.Column(
                            controls=[
                                ft.Container(
                                    content=ft.Row(
                                        controls=[
                                            ft.Text("Semantic Model & System Prompt", size=11, weight=ft.FontWeight.W_600, color=AppPalette.TEXT_MUTED),
                                            ft.IconButton(
                                                icon=ft.Icons.COPY_ALL,
                                                tooltip="Copy LLM Context",
                                                icon_size=16,
                                                on_click=self._copy_prompt_context,
                                            ),
                                        ],
                                        alignment=ft.MainAxisAlignment.SPACE_BETWEEN,
                                    ),
                                    padding=ft.Padding.symmetric(horizontal=8, vertical=4),
                                    bgcolor=AppPalette.BG_SURFACE,
                                    border=ft.Border.only(bottom=ft.BorderSide(1, AppPalette.BORDER_COLOR)),
                                ),
                                self.prompt_context_text,
                            ],
                            spacing=0,
                        ),
                        bgcolor=AppPalette.BG_CARD,
                        border=ft.Border.all(1, AppPalette.BORDER_COLOR),
                        border_radius=8,
                    ),
                ),
            ],
            expand=True,
        )

        # Right Panel: Protocol Input & Validation
        self.protocol_editor = CodeEditor(
            value=SAMPLE_ACTION_JSON,
            hint_text="Enter Action Protocol JSON here...",
            expand=True,
        )

        self.log_text = ft.Text("No execution logs yet.", size=12, color=AppPalette.TEXT_MUTED)
        self.validation_container = ft.Container(
            content=ft.Column(
                controls=[
                    ft.Text("Validation & Execution Log", size=12, weight=ft.FontWeight.BOLD, color=AppPalette.TEXT_PRIMARY),
                    self.log_text,
                ],
                spacing=8,
                scroll=ft.ScrollMode.AUTO,
            ),
            padding=12,
            bgcolor=AppPalette.BG_CARD,
            border=ft.Border.all(1, AppPalette.BORDER_COLOR),
            border_radius=8,
            height=130,
        )

        right_panel = ft.Column(
            controls=[
                self.protocol_editor,
                self.validation_container,
            ],
            spacing=8,
            expand=True,
        )

        # Split Body
        split_body = ft.Container(
            content=ft.Row(
                controls=[
                    ft.Container(content=left_tab_view, expand=4),
                    ft.Container(content=right_panel, expand=6),
                ],
                spacing=12,
                vertical_alignment=ft.CrossAxisAlignment.STRETCH,
            ),
            padding=ft.Padding.symmetric(horizontal=16),
            expand=True,
        )

        # Action Buttons
        self.open_file_btn = ActionButton(
            text="Open Workbook",
            icon=ft.Icons.FOLDER_OPEN_OUTLINED,
            color=AppPalette.PRIMARY,
            on_click=self._on_open_file_dialog,
        )

        self.execute_btn = ActionButton(
            text="Execute Protocol",
            icon=ft.Icons.PLAY_ARROW_ROUNDED,
            color=AppPalette.BRAND_EXCEL,
            tooltip="Validates and applies the JSON action protocol with automatic backup",
            on_click=self._on_execute_protocol,
        )

        self.backup_btn = ActionButton(
            text="View Backups",
            icon=ft.Icons.SHIELD_OUTLINED,
            color="#4A4F57",
            tooltip="View safe automated snapshot backups",
            on_click=self._on_view_backups,
        )

        action_row = ft.Container(
            content=ft.Row(
                controls=[
                    self.open_file_btn,
                    self.execute_btn,
                    self.backup_btn,
                ],
                spacing=10,
            ),
            padding=ft.Padding.symmetric(horizontal=16, vertical=10),
            bgcolor=AppPalette.BG_SURFACE,
            border=ft.Border.only(top=ft.BorderSide(1, AppPalette.BORDER_COLOR)),
        )

        main_layout = ft.Column(
            controls=[
                self.header,
                overview_row,
                split_body,
                action_row,
                self.status_bar,
            ],
            spacing=0,
            expand=True,
        )

        super().__init__(
            content=main_layout,
            expand=True,
            bgcolor=AppPalette.BG_DARK,
        )

    def _load_demo_workbook(self, e):
        demo_path = os.path.abspath("excel-copilot/Financial_Dashboard.xlsx")
        if not os.path.exists(demo_path):
            demo_path = os.path.abspath("Financial_Dashboard.xlsx")
        if os.path.exists(demo_path):
            self.load_workbook(demo_path)
        else:
            self.status_bar.set_status(f"Demo file not found at {demo_path}", level="warning")

    def _on_open_file_dialog(self, e):
        self._load_demo_workbook(e)

    def load_workbook(self, file_path: str):
        try:
            self.current_file_path = os.path.abspath(file_path)
            self.status_bar.set_status("Analyzing workbook structure…", level="info", loading=True)
            self.current_model = WorkbookAnalyzer.analyze(self.current_file_path)

            # Update Metric cards
            self.metric_file.content.controls[1].controls[1].value = os.path.basename(self.current_file_path)
            self.metric_sheets.content.controls[1].controls[1].value = str(len(self.current_model.worksheets))

            total_tables = sum(len(s.tables) for s in self.current_model.worksheets)
            self.metric_tables.content.controls[1].controls[1].value = str(total_tables)

            total_formulas = sum(s.formulas_count for s in self.current_model.worksheets)
            self.metric_formulas.content.controls[1].controls[1].value = str(total_formulas)

            # Populate sheets list
            self.sheets_list.controls.clear()
            for sheet in self.current_model.worksheets:
                card = ft.Container(
                    content=ft.Column(
                        controls=[
                            ft.Row(
                                controls=[
                                    ft.Icon(ft.Icons.GRID_ON, size=16, color=AppPalette.BRAND_EXCEL),
                                    ft.Text(sheet.name, size=13, weight=ft.FontWeight.BOLD, color=AppPalette.TEXT_PRIMARY),
                                    ft.Container(
                                        content=ft.Text(f"{sheet.max_row}x{sheet.max_column}", size=10, color=AppPalette.TEXT_MUTED),
                                        padding=ft.Padding.symmetric(horizontal=6, vertical=2),
                                        border_radius=4,
                                        bgcolor=AppPalette.BG_INPUT,
                                    ),
                                ],
                                spacing=8,
                            ),
                            ft.Text(f"Tables: {len(sheet.tables)}  |  Charts: {len(sheet.charts)}  |  Formulas: {sheet.formulas_count}", size=11, color=AppPalette.TEXT_SECONDARY),
                        ],
                        spacing=4,
                    ),
                    padding=10,
                    border_radius=6,
                    bgcolor=AppPalette.BG_CARD_HOVER,
                    border=ft.Border.all(1, AppPalette.BORDER_COLOR),
                )
                self.sheets_list.controls.append(card)

            # Generate LLM Context
            context_str = f"# EXCEL WORKBOOK SEMANTIC CONTEXT\nFile: {os.path.basename(self.current_file_path)}\n\n"
            for s in self.current_model.worksheets:
                context_str += f"## Sheet: {s.name} (Rows: {s.max_row}, Cols: {s.max_column})\n"
                if s.tables:
                    context_str += f"- Tables: {', '.join(t.name for t in s.tables)}\n"
                if s.formulas_count:
                    context_str += f"- Formulas Count: {s.formulas_count}\n"
            self.prompt_context_text.value = context_str

            self.status_bar.set_status(f"✓ Loaded '{os.path.basename(self.current_file_path)}' successfully", level="success")
            self.update()
        except Exception as err:
            self.status_bar.set_status(f"Failed to load workbook: {err}", level="error")

    def _copy_prompt_context(self, e):
        self.page.set_clipboard(self.prompt_context_text.value)
        self.status_bar.set_status("Copied LLM prompt context to clipboard", level="info")

    def _on_execute_protocol(self, e):
        if not self.current_file_path:
            self.status_bar.set_status("Please open or load a workbook first", level="warning")
            return

        json_text = self.protocol_editor.get_value().strip()
        if not json_text:
            self.status_bar.set_status("Action Protocol JSON is empty", level="warning")
            return

        self.status_bar.set_status("Validating and executing Action Protocol…", level="info", loading=True)

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
                    self.status_bar.set_status(f"✓ Protocol executed successfully! ({res.actions_executed} actions)", level="success")
                else:
                    log_lines.append(f"❌ {res.status.value}")
                    for err in res.errors:
                        log_lines.append(f"• Error: {err}")
                    self.status_bar.set_status(f"Execution finished with status: {res.status.value}", level="warning")

                for w in res.warnings:
                    log_lines.append(f"⚠ Warning: {w}")

                self.log_text.value = "\n".join(log_lines)
                self.log_text.color = AppPalette.SUCCESS if res.status == ExecutionStatus.SUCCESS else AppPalette.WARNING
                self.validation_container.update()
            except Exception as err:
                self.log_text.value = f"Parse/Execution Error:\n{str(err)}"
                self.log_text.color = AppPalette.ERROR
                self.validation_container.update()
                self.status_bar.set_status(f"Execution error: {err}", level="error")

        threading.Thread(target=worker, daemon=True).start()

    def _on_view_backups(self, e):
        backups = self.executor.backup_manager.list_backups()
        content_items = [ft.Text("Available Automatic Backups:", weight=ft.FontWeight.BOLD)]
        if not backups:
            content_items.append(ft.Text("No backup snapshots found yet.", color=AppPalette.TEXT_MUTED))
        else:
            for b in backups:
                content_items.append(
                    ft.Text(f"• {b['timestamp']} — {os.path.basename(b.get('source_path', ''))} ({b['size_bytes']} bytes)")
                )

        dialog = ft.AlertDialog(
            title=ft.Text("Automated Backups", size=16, weight=ft.FontWeight.BOLD),
            content=ft.Container(
                content=ft.Column(controls=content_items, spacing=6, scroll=ft.ScrollMode.AUTO),
                width=500,
                height=300,
            ),
            actions=[
                ft.TextButton("Close", on_click=lambda ev: self._close_dialog(dialog))
            ],
        )
        self.app_page.overlay.append(dialog)
        dialog.open = True
        self.app_page.update()

    def _close_dialog(self, dialog):
        dialog.open = False
        self.app_page.update()

