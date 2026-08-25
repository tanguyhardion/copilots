"""
CV Copilot View: JSON CV Editor, Deterministic DQ Engine Audit, and Europass Word (.docx) Generator.
"""

import os
import json
import threading
import flet as ft
from typing import Optional, Dict, Any, List

from copilots_app.core.theme import AppPalette
from copilots_app.ui.components import AppHeader, StatusBar, CodeEditor, ActionButton, MetricCard
from copilots_app.ui.prompt_dialog import open_prompt_dialog
from copilots_app.services.cv import load_sample_cv, run_dq_audit, generate_cv


class CVCopilotView(ft.Container):
    """Unified CV Copilot interface in Flet."""

    def __init__(self, page: ft.Page):
        self.app_page = page
        self.current_cv_data: Dict[str, Any] = load_sample_cv()

        self.header = AppHeader(
            title="CV Copilot",
            subtitle="Deterministic Data Quality engine, Europass profile validation, and formatted Word .docx generator",
            icon_path="icons/cv.png",
            badge_text="DQ Engine + docx",
            badge_color=AppPalette.BRAND_CV,
            actions=[
                ft.TextButton(
                    "System Prompt",
                    icon=ft.Icons.PSYCHOLOGY_OUTLINED,
                    on_click=lambda _: open_prompt_dialog(
                        self.app_page,
                        "cv",
                        on_status_change=lambda msg, lvl: self.status_bar.set_status(msg, level=lvl),
                    ),
                ),
                ft.TextButton(
                    "Reset Sample CV",
                    icon=ft.Icons.REFRESH,
                    on_click=self._load_default_sample,
                ),
            ],
        )

        self.status_bar = StatusBar(default_text="Ready — edit CV JSON and run Data Quality audit or build .docx")

        # Metric Cards
        self.metric_name = MetricCard(title="Candidate", value="Alex Morgan", icon_name=ft.Icons.PERSON_OUTLINE, color=AppPalette.BRAND_CV)
        self.metric_exp = MetricCard(title="Work Experience", value="3 Positions", icon_name=ft.Icons.WORK_OUTLINE, color=AppPalette.PRIMARY)
        self.metric_edu = MetricCard(title="Education & Certs", value="2 Degrees", icon_name=ft.Icons.SCHOOL_OUTLINED, color=AppPalette.INFO)
        self.metric_dq = MetricCard(title="DQ Status", value="0 Issues", icon_name=ft.Icons.VERIFIED_OUTLINED, color=AppPalette.SUCCESS)

        metrics_row = ft.Container(
            content=ft.Row(
                controls=[
                    self.metric_name,
                    self.metric_exp,
                    self.metric_edu,
                    self.metric_dq,
                ],
                spacing=10,
            ),
            padding=ft.Padding.symmetric(horizontal=16, vertical=12),
        )

        # JSON Editor (Left)
        sample_json_str = json.dumps(self.current_cv_data, indent=2) if self.current_cv_data else "{}"
        self.json_editor = CodeEditor(
            value=sample_json_str,
            hint_text="Enter structured CV JSON payload...",
            expand=True,
        )

        # DQ Issues List (Right)
        self.flags_list = ft.ListView(spacing=6, expand=True, padding=8)
        self.dq_panel = ft.Container(
            content=ft.Column(
                controls=[
                    ft.Container(
                        content=ft.Row(
                            controls=[
                                ft.Text("Data Quality & Compliance Audit", size=12, weight=ft.FontWeight.BOLD, color=AppPalette.TEXT_PRIMARY),
                                ft.Icon(ft.Icons.RULE_FOLDER_OUTLINED, size=16, color=AppPalette.TEXT_MUTED),
                            ],
                            alignment=ft.MainAxisAlignment.SPACE_BETWEEN,
                        ),
                        padding=ft.Padding.symmetric(horizontal=12, vertical=8),
                        bgcolor=AppPalette.BG_SURFACE,
                        border=ft.Border.only(bottom=ft.BorderSide(1, AppPalette.BORDER_COLOR)),
                    ),
                    self.flags_list,
                ],
                spacing=0,
                expand=True,
            ),
            bgcolor=AppPalette.BG_CARD,
            border=ft.Border.all(1, AppPalette.BORDER_COLOR),
            border_radius=8,
            expand=True,
        )

        # Split Workspace
        split_workspace = ft.Container(
            content=ft.Row(
                controls=[
                    ft.Container(content=self.json_editor, expand=6),
                    ft.Container(content=self.dq_panel, expand=4),
                ],
                spacing=12,
                vertical_alignment=ft.CrossAxisAlignment.STRETCH,
            ),
            padding=ft.Padding.symmetric(horizontal=16),
            expand=True,
        )

        # Action Bar
        self.audit_btn = ActionButton(
            text="Run DQ Audit",
            icon=ft.Icons.RULE_ROUNDED,
            color=AppPalette.BRAND_CV,
            tooltip="Executes deterministic data quality & completeness validation rules",
            on_click=self._on_run_audit,
        )

        self.generate_btn = ActionButton(
            text="Generate Word CV (.docx)",
            icon=ft.Icons.DOWNLOAD_ROUNDED,
            color=AppPalette.SUCCESS,
            tooltip="Generates formatted Europass standard CV in Microsoft Word docx format",
            on_click=self._on_generate_cv,
        )

        self.open_doc_btn = ActionButton(
            text="Generate & Open in Word",
            icon=ft.Icons.OPEN_IN_NEW,
            color=AppPalette.PRIMARY,
            tooltip="Generates and automatically launches the Word CV",
            on_click=lambda e: self._on_generate_cv(e, auto_open=True),
        )

        action_row = ft.Container(
            content=ft.Row(
                controls=[
                    self.audit_btn,
                    self.generate_btn,
                    self.open_doc_btn,
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
                metrics_row,
                split_workspace,
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

        # Initial summary update
        self._update_summary_metrics(self.current_cv_data)

    def _load_default_sample(self, e):
        self.current_cv_data = load_sample_cv()
        self.json_editor.set_value(json.dumps(self.current_cv_data, indent=2))
        self._update_summary_metrics(self.current_cv_data)
        self.status_bar.set_status("Loaded sample CV dataset", level="info")

    def _update_summary_metrics(self, data: Dict[str, Any]):
        try:
            pi = data.get("personal_info", {})
            name = f"{pi.get('first_name', '')} {pi.get('last_name', '')}".strip() or "Candidate"
            self.metric_name.content.controls[1].controls[1].value = name

            work_count = len(data.get("work_experience", []))
            self.metric_exp.content.controls[1].controls[1].value = f"{work_count} Positions"

            edu_count = len(data.get("education", []))
            certs_count = len(data.get("personal_skills", {}).get("certifications", []))
            self.metric_edu.content.controls[1].controls[1].value = f"{edu_count} Edu / {certs_count} Certs"

            self.update()
        except Exception:
            pass

    def _parse_editor_json(self) -> Optional[Dict[str, Any]]:
        text = self.json_editor.get_value().strip()
        if not text:
            self.status_bar.set_status("JSON content is empty", level="warning")
            return None
        try:
            data = json.loads(text)
            self._update_summary_metrics(data)
            return data
        except Exception as err:
            self.status_bar.set_status(f"JSON Syntax Error: {err}", level="error")
            return None

    def _on_run_audit(self, e):
        data = self._parse_editor_json()
        if not data:
            return

        self.status_bar.set_status("Running deterministic DQ rule checks…", level="info", loading=True)
        flags = run_dq_audit(data)

        self.flags_list.controls.clear()
        errors_count = sum(1 for f in flags if f.get("severity") == "ERROR")
        warnings_count = sum(1 for f in flags if f.get("severity") == "WARNING")

        if not flags:
            self.flags_list.controls.append(
                ft.Container(
                    content=ft.Row(
                        controls=[
                            ft.Icon(ft.Icons.CHECK_CIRCLE, color=AppPalette.SUCCESS, size=20),
                            ft.Text("All Data Quality rules passed! 100% compliant.", size=12, color=AppPalette.SUCCESS),
                        ],
                        spacing=8,
                    ),
                    padding=16,
                )
            )
            self.metric_dq.content.controls[1].controls[1].value = "✓ 100% Pass"
            self.metric_dq.content.controls[1].controls[1].color = AppPalette.SUCCESS
            self.status_bar.set_status("✓ Data Quality audit passed with zero flags", level="success")
        else:
            for flag in flags:
                sev = flag.get("severity", "WARNING")
                is_err = sev == "ERROR"
                color = AppPalette.ERROR if is_err else AppPalette.WARNING
                icon = ft.Icons.ERROR_OUTLINE if is_err else ft.Icons.WARNING_AMBER_ROUNDED

                item = ft.Container(
                    content=ft.Row(
                        controls=[
                            ft.Icon(icon, color=color, size=16),
                            ft.Column(
                                controls=[
                                    ft.Row(
                                        controls=[
                                            ft.Text(flag.get("rule_code", "FLAG"), size=11, weight=ft.FontWeight.BOLD, color=color),
                                            ft.Text(f"[{flag.get('section', 'General')}]", size=10, color=AppPalette.TEXT_MUTED),
                                        ],
                                        spacing=6,
                                    ),
                                    ft.Text(flag.get("message", ""), size=12, color=AppPalette.TEXT_PRIMARY),
                                ],
                                spacing=2,
                                expand=True,
                            ),
                        ],
                        spacing=10,
                        vertical_alignment=ft.CrossAxisAlignment.START,
                    ),
                    padding=10,
                    border_radius=6,
                    bgcolor=AppPalette.BG_CARD_HOVER,
                    border=ft.Border.all(1, color),
                )
                self.flags_list.controls.append(item)

            self.metric_dq.content.controls[1].controls[1].value = f"{errors_count} Err / {warnings_count} Warn"
            self.metric_dq.content.controls[1].controls[1].color = AppPalette.ERROR if errors_count else AppPalette.WARNING
            self.status_bar.set_status(f"DQ Audit complete: {errors_count} errors, {warnings_count} warnings", level="warning" if errors_count else "info")

        self.flags_list.update()
        self.metric_dq.update()

    def _on_generate_cv(self, e, auto_open: bool = False):
        data = self._parse_editor_json()
        if not data:
            return

        out_path = os.path.abspath("Generated_Europass_CV.docx")
        self.status_bar.set_status("Generating formatted Word CV (.docx)…", level="info", loading=True)

        def worker():
            try:
                generate_cv(data, out_path)
                self.status_bar.set_status(f"✓ CV saved successfully to '{out_path}'", level="success")
                if auto_open and os.path.exists(out_path):
                    os.startfile(out_path)
            except Exception as err:
                self.status_bar.set_status(f"CV generation failed: {err}", level="error")

        threading.Thread(target=worker, daemon=True).start()
