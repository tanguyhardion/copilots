"""
CV Copilot View in PySide6: JSON CV Editor, Deterministic DQ Engine Audit, and Europass Word (.docx) Generator.
"""

import os
import json
import threading
from typing import Optional, Dict, Any, List
from PySide6.QtWidgets import (
    QWidget,
    QVBoxLayout,
    QHBoxLayout,
    QFrame,
    QPushButton,
    QLabel,
    QScrollArea,
)
from PySide6.QtCore import Qt, Signal, QObject
from PySide6.QtGui import QFont

from copilots_app.core.theme import AppPalette
from copilots_app.ui.components import AppHeader, StatusBar, CodeEditor, ActionButton, MetricCard
from copilots_app.ui.prompt_dialog import open_prompt_dialog
from copilots_app.services.cv import load_sample_cv, run_dq_audit, generate_cv


class WorkerSignals(QObject):
    status = Signal(str, str, bool)
    finished = Signal()


class CVCopilotView(QWidget):
    """Unified CV Copilot interface in PySide6."""

    def __init__(self, parent: Optional[QWidget] = None):
        super().__init__(parent)
        self.current_cv_data: Dict[str, Any] = load_sample_cv()

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
            "cv",
            on_status_change=lambda msg, lvl: self.status_bar.set_status(msg, level=lvl),
            parent=self,
        ))

        reset_btn = QPushButton("Reset Sample CV")
        reset_btn.setFont(QFont("Segoe UI", 9))
        reset_btn.setStyleSheet(f"""
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
        reset_btn.clicked.connect(self._load_default_sample)

        self.header = AppHeader(
            title="CV Copilot",
            subtitle="Deterministic Data Quality engine, Europass profile validation, and formatted Word .docx generator",
            icon_path="icons/cv.png",
            badge_text="DQ Engine + docx",
            badge_color=AppPalette.BRAND_CV,
            actions=[prompt_btn, reset_btn],
            parent=self,
        )
        layout.addWidget(self.header)

        # Metric Cards Row
        metrics_container = QWidget()
        m_layout = QHBoxLayout(metrics_container)
        m_layout.setContentsMargins(16, 12, 16, 12)
        m_layout.setSpacing(10)

        self.metric_name = MetricCard(title="Candidate", value="Alex Morgan", icon_text="👤", color=AppPalette.BRAND_CV)
        self.metric_exp = MetricCard(title="Work Experience", value="3 Positions", icon_text="💼", color=AppPalette.PRIMARY)
        self.metric_edu = MetricCard(title="Education & Certs", value="2 Degrees", icon_text="🎓", color=AppPalette.INFO)
        self.metric_dq = MetricCard(title="DQ Status", value="0 Issues", icon_text="✓", color=AppPalette.SUCCESS)

        m_layout.addWidget(self.metric_name)
        m_layout.addWidget(self.metric_exp)
        m_layout.addWidget(self.metric_edu)
        m_layout.addWidget(self.metric_dq)
        layout.addWidget(metrics_container)

        # Split Workspace (Left: JSON Editor, Right: DQ Audit list)
        split_widget = QWidget()
        split_layout = QHBoxLayout(split_widget)
        split_layout.setContentsMargins(16, 0, 16, 16)
        split_layout.setSpacing(12)

        sample_json_str = json.dumps(self.current_cv_data, indent=2) if self.current_cv_data else "{}"
        self.json_editor = CodeEditor(
            value=sample_json_str,
            hint_text="Enter structured CV JSON payload...",
            parent=self,
        )
        split_layout.addWidget(self.json_editor, 6)

        # DQ Panel
        dq_panel = QFrame()
        dq_panel.setStyleSheet(f"""
            background-color: {AppPalette.BG_CARD};
            border: 1px solid {AppPalette.BORDER_COLOR};
            border-radius: 8px;
        """)
        dq_layout = QVBoxLayout(dq_panel)
        dq_layout.setContentsMargins(0, 0, 0, 0)
        dq_layout.setSpacing(0)

        # DQ Header
        dq_h = QFrame()
        dq_h.setStyleSheet(f"""
            background-color: {AppPalette.BG_SURFACE};
            border-bottom: 1px solid {AppPalette.BORDER_COLOR};
            border-top-left-radius: 8px;
            border-top-right-radius: 8px;
        """)
        dq_h_layout = QHBoxLayout(dq_h)
        dq_h_layout.setContentsMargins(14, 10, 14, 10)
        dq_title = QLabel("Data Quality & Compliance Audit")
        dq_title.setFont(QFont("Segoe UI", 9, QFont.Bold))
        dq_title.setStyleSheet(f"color: {AppPalette.TEXT_PRIMARY};")
        dq_h_layout.addWidget(dq_title)
        dq_layout.addWidget(dq_h)

        # Flags scroll area
        self.flags_scroll = QScrollArea()
        self.flags_scroll.setWidgetResizable(True)
        self.flags_scroll.setStyleSheet("border: none; background: transparent;")
        self.flags_container = QWidget()
        self.flags_layout = QVBoxLayout(self.flags_container)
        self.flags_layout.setContentsMargins(10, 10, 10, 10)
        self.flags_layout.setSpacing(8)
        self.flags_layout.addStretch()
        self.flags_scroll.setWidget(self.flags_container)
        dq_layout.addWidget(self.flags_scroll)

        split_layout.addWidget(dq_panel, 4)
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

        self.audit_btn = ActionButton(
            text="Run DQ Audit",
            color=AppPalette.BRAND_CV,
            tooltip="Executes deterministic data quality & completeness validation rules",
        )
        self.audit_btn.clicked.connect(self._on_run_audit)
        act_layout.addWidget(self.audit_btn)

        self.generate_btn = ActionButton(
            text="Generate Word CV (.docx)",
            color=AppPalette.SUCCESS,
            tooltip="Generates formatted Europass standard CV in Microsoft Word docx format",
        )
        self.generate_btn.clicked.connect(lambda: self._on_generate_cv(auto_open=False))
        act_layout.addWidget(self.generate_btn)

        self.open_doc_btn = ActionButton(
            text="Generate & Open in Word",
            color=AppPalette.PRIMARY,
            tooltip="Generates and automatically launches the Word CV",
        )
        self.open_doc_btn.clicked.connect(lambda: self._on_generate_cv(auto_open=True))
        act_layout.addWidget(self.open_doc_btn)

        act_layout.addStretch()
        layout.addWidget(action_bar)

        # Status Bar
        self.status_bar = StatusBar(default_text="Ready — edit CV JSON and run Data Quality audit or build .docx", parent=self)
        layout.addWidget(self.status_bar)

        # Initial metrics
        self._update_summary_metrics(self.current_cv_data)

    def _load_default_sample(self):
        self.current_cv_data = load_sample_cv()
        self.json_editor.set_value(json.dumps(self.current_cv_data, indent=2))
        self._update_summary_metrics(self.current_cv_data)
        self.status_bar.set_status("Loaded sample CV dataset", level="info")

    def _update_summary_metrics(self, data: Dict[str, Any]):
        try:
            pi = data.get("personal_info", {})
            name = f"{pi.get('first_name', '')} {pi.get('last_name', '')}".strip() or "Candidate"
            self.metric_name.set_value(name)

            work_count = len(data.get("work_experience", []))
            self.metric_exp.set_value(f"{work_count} Positions")

            edu_count = len(data.get("education", []))
            certs_count = len(data.get("personal_skills", {}).get("certifications", []))
            self.metric_edu.set_value(f"{edu_count} Edu / {certs_count} Certs")
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

    def _on_run_audit(self):
        data = self._parse_editor_json()
        if not data:
            return

        self.status_bar.set_status("Running deterministic DQ rule checks…", level="info", loading=True)
        flags = run_dq_audit(data)

        # Clear previous flags
        while self.flags_layout.count() > 1:
            child = self.flags_layout.takeAt(0)
            if child.widget():
                child.widget().deleteLater()

        errors_count = sum(1 for f in flags if f.get("severity") == "ERROR")
        warnings_count = sum(1 for f in flags if f.get("severity") == "WARNING")

        if not flags:
            success_card = QFrame()
            success_card.setStyleSheet(f"""
                background-color: {AppPalette.BG_CARD_HOVER};
                border: 1px solid {AppPalette.SUCCESS};
                border-radius: 6px;
                padding: 12px;
            """)
            s_layout = QHBoxLayout(success_card)
            s_label = QLabel("✓ All Data Quality rules passed! 100% compliant.")
            s_label.setFont(QFont("Segoe UI", 9, QFont.Bold))
            s_label.setStyleSheet(f"color: {AppPalette.SUCCESS};")
            s_layout.addWidget(s_label)
            self.flags_layout.insertWidget(0, success_card)

            self.metric_dq.set_value("✓ 100% Pass", color=AppPalette.SUCCESS)
            self.status_bar.set_status("✓ Data Quality audit passed with zero flags", level="success")
        else:
            for flag in flags:
                sev = flag.get("severity", "WARNING")
                is_err = (sev == "ERROR")
                color = AppPalette.ERROR if is_err else AppPalette.WARNING

                flag_card = QFrame()
                flag_card.setStyleSheet(f"""
                    background-color: {AppPalette.BG_CARD_HOVER};
                    border: 1px solid {color};
                    border-radius: 6px;
                """)
                f_layout = QVBoxLayout(flag_card)
                f_layout.setContentsMargins(10, 8, 10, 8)
                f_layout.setSpacing(2)

                top_r = QHBoxLayout()
                r_code = QLabel(flag.get("rule_code", "FLAG"))
                r_code.setFont(QFont("Segoe UI", 8, QFont.Bold))
                r_code.setStyleSheet(f"color: {color};")
                top_r.addWidget(r_code)

                r_sec = QLabel(f"[{flag.get('section', 'General')}]")
                r_sec.setFont(QFont("Segoe UI", 8))
                r_sec.setStyleSheet(f"color: {AppPalette.TEXT_MUTED};")
                top_r.addWidget(r_sec)
                top_r.addStretch()
                f_layout.addLayout(top_r)

                msg_lbl = QLabel(flag.get("message", ""))
                msg_lbl.setFont(QFont("Segoe UI", 9))
                msg_lbl.setStyleSheet(f"color: {AppPalette.TEXT_PRIMARY};")
                msg_lbl.setWordWrap(True)
                f_layout.addWidget(msg_lbl)

                self.flags_layout.insertWidget(self.flags_layout.count() - 1, flag_card)

            self.metric_dq.set_value(f"{errors_count} Err / {warnings_count} Warn", color=AppPalette.ERROR if errors_count else AppPalette.WARNING)
            self.status_bar.set_status(f"DQ Audit complete: {errors_count} errors, {warnings_count} warnings", level="warning" if errors_count else "info")

    def _on_generate_cv(self, auto_open: bool = False):
        data = self._parse_editor_json()
        if not data:
            return

        out_path = os.path.abspath("Generated_Europass_CV.docx")
        self.status_bar.set_status("Generating formatted Word CV (.docx)…", level="info", loading=True)

        signals = WorkerSignals()
        signals.status.connect(self.status_bar.set_status)

        def worker():
            try:
                generate_cv(data, out_path)
                signals.status.emit(f"✓ CV saved successfully to '{out_path}'", "success", False)
                if auto_open and os.path.exists(out_path):
                    os.startfile(out_path)
            except Exception as err:
                signals.status.emit(f"CV generation failed: {err}", "error", False)

        threading.Thread(target=worker, daemon=True).start()
