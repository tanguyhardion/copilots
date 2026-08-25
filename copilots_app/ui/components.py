"""
Reusable UI Components in PySide6: AppHeader, StatusBar, CodeEditor, MetricCard, and ActionButton.
"""

import os
from typing import Optional, Dict, Any, List, Callable
from PySide6.QtWidgets import (
    QWidget,
    QHBoxLayout,
    QVBoxLayout,
    QLabel,
    QPushButton,
    QPlainTextEdit,
    QComboBox,
    QFrame,
    QProgressBar,
    QApplication,
    QGraphicsDropShadowEffect,
)
from PySide6.QtCore import Qt, Signal, QSize
from PySide6.QtGui import QIcon, QPixmap, QFont, QColor, QPainter, QBrush, QPen

from copilots_app.core.theme import AppPalette, get_asset_path


class AppHeader(QFrame):
    """Modern header banner for Copilot views with title, subtitle, icon, badge, and action buttons."""

    def __init__(
        self,
        title: str,
        subtitle: str,
        icon_path: Optional[str] = None,
        badge_text: Optional[str] = None,
        badge_color: str = AppPalette.PRIMARY,
        actions: Optional[List[QPushButton]] = None,
        parent: Optional[QWidget] = None,
    ):
        super().__init__(parent)
        self.setObjectName("AppHeader")
        self.setStyleSheet(f"""
            QFrame#AppHeader {{
                background-color: {AppPalette.BG_SURFACE};
                border-bottom: 1px solid {AppPalette.BORDER_COLOR};
            }}
        """)

        layout = QHBoxLayout(self)
        layout.setContentsMargins(20, 14, 20, 14)
        layout.setSpacing(14)

        # Left: Icon + Title/Subtitle + Badge
        left_layout = QHBoxLayout()
        left_layout.setSpacing(12)

        if icon_path:
            resolved_icon = get_asset_path(icon_path)
            icon_label = QLabel()
            if os.path.exists(resolved_icon):
                pixmap = QPixmap(resolved_icon).scaled(
                    32, 32, Qt.KeepAspectRatio, Qt.SmoothTransformation
                )
                icon_label.setPixmap(pixmap)
            icon_label.setStyleSheet(f"""
                background-color: {AppPalette.BG_CARD};
                border: 1px solid {AppPalette.BORDER_COLOR};
                border-radius: 18px;
                padding: 4px;
            """)
            left_layout.addWidget(icon_label)

        text_layout = QVBoxLayout()
        text_layout.setSpacing(3)
        text_layout.setContentsMargins(0, 0, 0, 0)

        title_row = QHBoxLayout()
        title_row.setSpacing(10)

        title_label = QLabel(title)
        title_label.setFont(QFont("Segoe UI", 13, QFont.Bold))
        title_label.setStyleSheet(f"color: {AppPalette.TEXT_PRIMARY};")
        title_row.addWidget(title_label)

        if badge_text:
            badge_label = QLabel(badge_text)
            badge_label.setFont(QFont("Segoe UI", 8, QFont.Bold))
            badge_label.setStyleSheet(f"""
                background-color: {badge_color};
                color: #FFFFFF;
                border-radius: 10px;
                padding: 2px 8px;
            """)
            title_row.addWidget(badge_label)

        title_row.addStretch()
        text_layout.addLayout(title_row)

        subtitle_label = QLabel(subtitle)
        subtitle_label.setFont(QFont("Segoe UI", 9))
        subtitle_label.setStyleSheet(f"color: {AppPalette.TEXT_SECONDARY};")
        text_layout.addWidget(subtitle_label)

        left_layout.addLayout(text_layout)
        layout.addLayout(left_layout)
        layout.addStretch()

        # Right: Action buttons
        if actions:
            actions_layout = QHBoxLayout()
            actions_layout.setSpacing(8)
            for btn in actions:
                actions_layout.addWidget(btn)
            layout.addLayout(actions_layout)


class StatusBar(QFrame):
    """Async status indicator bar with levels (info, success, warning, error) and indicator."""

    def __init__(self, default_text: str = "Ready", parent: Optional[QWidget] = None):
        super().__init__(parent)
        self.setObjectName("StatusBar")
        self.setStyleSheet(f"""
            QFrame#StatusBar {{
                background-color: {AppPalette.BG_SURFACE};
                border-top: 1px solid {AppPalette.BORDER_COLOR};
            }}
        """)

        layout = QHBoxLayout(self)
        layout.setContentsMargins(16, 7, 16, 7)
        layout.setSpacing(8)

        self.indicator = QLabel("●")
        self.indicator.setFont(QFont("Segoe UI", 10, QFont.Bold))
        self.indicator.setStyleSheet(f"color: {AppPalette.TEXT_MUTED};")
        layout.addWidget(self.indicator)

        self.spinner = QProgressBar()
        self.spinner.setRange(0, 0)  # Indeterminate mode
        self.spinner.setFixedSize(60, 10)
        self.spinner.setTextVisible(False)
        self.spinner.setStyleSheet(f"""
            QProgressBar {{
                background-color: {AppPalette.BG_INPUT};
                border: 1px solid {AppPalette.BORDER_COLOR};
                border-radius: 4px;
            }}
            QProgressBar::chunk {{
                background-color: {AppPalette.PRIMARY};
                border-radius: 3px;
            }}
        """)
        self.spinner.setVisible(False)
        layout.addWidget(self.spinner)

        self.text_label = QLabel(default_text)
        self.text_label.setFont(QFont("Segoe UI", 9))
        self.text_label.setStyleSheet(f"color: {AppPalette.TEXT_SECONDARY};")
        layout.addWidget(self.text_label)
        layout.addStretch()

    def set_status(self, message: str, level: str = "info", loading: bool = False):
        color_map = {
            "info": AppPalette.TEXT_SECONDARY,
            "success": AppPalette.SUCCESS,
            "warning": AppPalette.WARNING,
            "error": AppPalette.ERROR,
        }
        color = color_map.get(level, AppPalette.TEXT_SECONDARY)

        self.spinner.setVisible(loading)
        self.indicator.setVisible(not loading)
        self.indicator.setStyleSheet(f"color: {color};")
        self.text_label.setText(message)
        self.text_label.setStyleSheet(f"color: {color};")
        QApplication.processEvents()


class CodeEditor(QFrame):
    """Monospace DSL & JSON code editor with toolbar (samples dropdown, Copy, Clear)."""

    def __init__(
        self,
        value: str = "",
        hint_text: str = "Paste or enter DSL / Code here...",
        samples: Optional[Dict[str, str]] = None,
        on_sample_selected: Optional[Callable[[str], None]] = None,
        parent: Optional[QWidget] = None,
    ):
        super().__init__(parent)
        self.samples = samples or {}
        self.on_sample_selected = on_sample_selected

        self.setObjectName("CodeEditor")
        self.setStyleSheet(f"""
            QFrame#CodeEditor {{
                background-color: {AppPalette.BG_INPUT};
                border: 1px solid {AppPalette.BORDER_COLOR};
                border-radius: 8px;
            }}
        """)

        layout = QVBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(0)

        # Toolbar
        toolbar = QFrame()
        toolbar.setStyleSheet(f"""
            background-color: {AppPalette.BG_SURFACE};
            border-bottom: 1px solid {AppPalette.BORDER_COLOR};
            border-top-left-radius: 8px;
            border-top-right-radius: 8px;
            padding: 2px;
        """)
        tb_layout = QHBoxLayout(toolbar)
        tb_layout.setContentsMargins(12, 6, 12, 6)
        tb_layout.setSpacing(10)

        tb_title = QLabel("DSL / Code Editor")
        tb_title.setFont(QFont("Segoe UI", 9, QFont.Bold))
        tb_title.setStyleSheet(f"color: {AppPalette.TEXT_MUTED};")
        tb_layout.addWidget(tb_title)

        tb_layout.addStretch()

        if self.samples:
            self.sample_combo = QComboBox()
            self.sample_combo.addItem("Load sample template...")
            for key in self.samples.keys():
                self.sample_combo.addItem(key)
            self.sample_combo.currentIndexChanged.connect(self._handle_sample_change)
            tb_layout.addWidget(self.sample_combo)

        copy_btn = QPushButton("Copy All")
        copy_btn.setFont(QFont("Segoe UI", 9))
        copy_btn.setStyleSheet(f"""
            QPushButton {{
                background-color: {AppPalette.BG_CARD};
                color: {AppPalette.TEXT_SECONDARY};
                border: 1px solid {AppPalette.BORDER_COLOR};
                border-radius: 5px;
                padding: 4px 10px;
            }}
            QPushButton:hover {{
                background-color: {AppPalette.BG_CARD_HOVER};
                color: {AppPalette.TEXT_PRIMARY};
                border-color: {AppPalette.BORDER_LIGHT};
            }}
        """)
        copy_btn.clicked.connect(self._copy_all)
        tb_layout.addWidget(copy_btn)

        clear_btn = QPushButton("Clear")
        clear_btn.setFont(QFont("Segoe UI", 9))
        clear_btn.setStyleSheet(f"""
            QPushButton {{
                background-color: {AppPalette.BG_CARD};
                color: {AppPalette.TEXT_SECONDARY};
                border: 1px solid {AppPalette.BORDER_COLOR};
                border-radius: 5px;
                padding: 4px 10px;
            }}
            QPushButton:hover {{
                background-color: {AppPalette.BG_CARD_HOVER};
                color: {AppPalette.ERROR};
                border-color: {AppPalette.ERROR};
            }}
        """)
        clear_btn.clicked.connect(self.clear)
        tb_layout.addWidget(clear_btn)

        layout.addWidget(toolbar)

        # Plain text edit
        self.text_edit = QPlainTextEdit()
        self.text_edit.setPlainText(value)
        self.text_edit.setPlaceholderText(hint_text)
        self.text_edit.setFont(QFont("Consolas", 10))
        self.text_edit.setStyleSheet(f"""
            QPlainTextEdit {{
                background-color: {AppPalette.BG_INPUT};
                color: {AppPalette.TEXT_PRIMARY};
                border: none;
                border-bottom-left-radius: 8px;
                border-bottom-right-radius: 8px;
                padding: 12px;
                selection-background-color: {AppPalette.PRIMARY};
                selection-color: #FFFFFF;
            }}
        """)
        layout.addWidget(self.text_edit)

    def _handle_sample_change(self, index: int):
        if index <= 0:
            return
        key = self.sample_combo.currentText()
        if key in self.samples:
            self.set_value(self.samples[key])
            if self.on_sample_selected:
                self.on_sample_selected(key)

    def _copy_all(self):
        text = self.get_value()
        QApplication.clipboard().setText(text)

    def clear(self):
        self.text_edit.clear()

    def get_value(self) -> str:
        return self.text_edit.toPlainText()

    def set_value(self, val: str):
        self.text_edit.setPlainText(val)


class MetricCard(QFrame):
    """Sleek KPI / Summary Metric Card with accent icon."""

    def __init__(
        self,
        title: str,
        value: str,
        icon_text: str = "◆",
        color: str = AppPalette.PRIMARY,
        subtitle: str = "",
        parent: Optional[QWidget] = None,
    ):
        super().__init__(parent)
        self.setObjectName("MetricCard")
        self.setStyleSheet(f"""
            QFrame#MetricCard {{
                background-color: {AppPalette.BG_CARD};
                border: 1px solid {AppPalette.BORDER_COLOR};
                border-radius: 8px;
            }}
        """)

        layout = QHBoxLayout(self)
        layout.setContentsMargins(14, 12, 14, 12)
        layout.setSpacing(12)

        icon_box = QLabel(icon_text)
        icon_box.setFixedSize(38, 38)
        icon_box.setAlignment(Qt.AlignCenter)
        icon_box.setFont(QFont("Segoe UI", 14, QFont.Bold))
        icon_box.setStyleSheet(f"""
            background-color: {AppPalette.BG_INPUT};
            color: {color};
            border: 1px solid {AppPalette.BORDER_COLOR};
            border-radius: 8px;
        """)
        layout.addWidget(icon_box)

        text_layout = QVBoxLayout()
        text_layout.setSpacing(2)
        text_layout.setContentsMargins(0, 0, 0, 0)

        title_label = QLabel(title)
        title_label.setFont(QFont("Segoe UI", 8, QFont.Bold))
        title_label.setStyleSheet(f"color: {AppPalette.TEXT_SECONDARY}; text-transform: uppercase;")
        text_layout.addWidget(title_label)

        self.value_label = QLabel(value)
        self.value_label.setFont(QFont("Segoe UI", 13, QFont.Bold))
        self.value_label.setStyleSheet(f"color: {AppPalette.TEXT_PRIMARY};")
        text_layout.addWidget(self.value_label)

        if subtitle:
            sub_label = QLabel(subtitle)
            sub_label.setFont(QFont("Segoe UI", 8))
            sub_label.setStyleSheet(f"color: {AppPalette.TEXT_MUTED};")
            text_layout.addWidget(sub_label)

        layout.addLayout(text_layout)
        layout.addStretch()

    def set_value(self, val: str, color: Optional[str] = None):
        self.value_label.setText(val)
        if color:
            self.value_label.setStyleSheet(f"color: {color};")


class ActionButton(QPushButton):
    """Styled action button with custom accent color."""

    def __init__(
        self,
        text: str,
        color: str = AppPalette.PRIMARY,
        tooltip: Optional[str] = None,
        parent: Optional[QWidget] = None,
    ):
        super().__init__(text, parent)
        if tooltip:
            self.setToolTip(tooltip)
        self.setFont(QFont("Segoe UI", 9, QFont.Bold))
        self.setCursor(Qt.PointingHandCursor)
        self.setStyleSheet(f"""
            QPushButton {{
                background-color: {color};
                color: #FFFFFF;
                border: none;
                border-radius: 6px;
                padding: 8px 16px;
                min-height: 20px;
            }}
            QPushButton:hover {{
                background-color: {color}EE;
                border: 1px solid #FFFFFF44;
            }}
            QPushButton:pressed {{
                background-color: {color}CC;
            }}
            QPushButton:disabled {{
                background-color: {AppPalette.BORDER_COLOR};
                color: {AppPalette.TEXT_MUTED};
            }}
        """)
