"""
Sidebar navigation menu in PySide6 with brand header, active route indicators, and settings footer.
"""

import os
from typing import Callable, Optional, Dict
from PySide6.QtWidgets import (
    QWidget,
    QVBoxLayout,
    QHBoxLayout,
    QLabel,
    QPushButton,
    QFrame,
)
from PySide6.QtCore import Qt, Signal
from PySide6.QtGui import QFont, QPixmap, QCursor

from copilots_app.core.theme import AppPalette, get_asset_path
from copilots_app.core.prompt_manager import PromptManager


class NavItem(QFrame):
    """Custom navigation sidebar item with icon, label, and active indicator line."""

    clicked = Signal(str)

    def __init__(
        self,
        route_id: str,
        title: str,
        icon_path: Optional[str] = None,
        badge_color: str = AppPalette.PRIMARY,
        is_selected: bool = False,
        parent: Optional[QWidget] = None,
    ):
        super().__init__(parent)
        self.route_id = route_id
        self.badge_color = badge_color
        self.is_selected = is_selected

        self.setFixedHeight(44)
        self.setCursor(Qt.PointingHandCursor)

        layout = QHBoxLayout(self)
        layout.setContentsMargins(10, 4, 12, 4)
        layout.setSpacing(10)

        # Left indicator bar
        self.indicator = QFrame()
        self.indicator.setFixedWidth(3)
        self.indicator.setFixedHeight(22)
        layout.addWidget(self.indicator)

        # Icon
        self.icon_label = QLabel()
        self.icon_label.setFixedSize(26, 26)
        if icon_path:
            resolved_icon = get_asset_path(icon_path)
            if os.path.exists(resolved_icon):
                pix = QPixmap(resolved_icon).scaled(24, 24, Qt.KeepAspectRatio, Qt.SmoothTransformation)
                self.icon_label.setPixmap(pix)
        layout.addWidget(self.icon_label)

        # Label
        self.label = QLabel(title)
        self.label.setFont(QFont("Segoe UI", 10))
        layout.addWidget(self.label)
        layout.addStretch()

        self._update_appearance()

    def set_selected(self, selected: bool):
        self.is_selected = selected
        self._update_appearance()

    def _update_appearance(self):
        if self.is_selected:
            self.setStyleSheet(f"""
                QFrame {{
                    background-color: {AppPalette.BG_CARD};
                    border-radius: 6px;
                }}
            """)
            self.indicator.setStyleSheet(f"""
                background-color: {self.badge_color};
                border-radius: 2px;
            """)
            self.label.setFont(QFont("Segoe UI", 10, QFont.Bold))
            self.label.setStyleSheet(f"color: {AppPalette.TEXT_PRIMARY};")
        else:
            self.setStyleSheet(f"""
                QFrame {{
                    background-color: transparent;
                    border-radius: 6px;
                }}
                QFrame:hover {{
                    background-color: {AppPalette.BG_CARD_HOVER};
                }}
            """)
            self.indicator.setStyleSheet("background-color: transparent;")
            self.label.setFont(QFont("Segoe UI", 10))
            self.label.setStyleSheet(f"color: {AppPalette.TEXT_SECONDARY};")

    def mousePressEvent(self, event):
        if event.button() == Qt.LeftButton:
            self.clicked.emit(self.route_id)
        super().mousePressEvent(event)


class AppSidebar(QFrame):
    """Main application sidebar containing brand headers, navigation items, and settings."""

    def __init__(
        self,
        current_route: str = "powerpoint",
        on_navigate: Optional[Callable[[str], None]] = None,
        parent: Optional[QWidget] = None,
    ):
        super().__init__(parent)
        self.on_navigate = on_navigate
        self.current_route = current_route

        self.setFixedWidth(240)
        self.setObjectName("AppSidebar")
        self.setStyleSheet(f"""
            QFrame#AppSidebar {{
                background-color: {AppPalette.BG_SURFACE};
                border-right: 1px solid {AppPalette.BORDER_COLOR};
            }}
        """)

        layout = QVBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(0)

        # Brand Header
        header_frame = QFrame()
        header_frame.setStyleSheet(f"""
            border-bottom: 1px solid {AppPalette.BORDER_COLOR};
            padding: 4px;
        """)
        h_layout = QHBoxLayout(header_frame)
        h_layout.setContentsMargins(16, 16, 16, 16)
        h_layout.setSpacing(12)

        app_icon_label = QLabel()
        app_icon_label.setFixedSize(32, 32)
        app_icon_path = get_asset_path("icons/copilots.png")
        if os.path.exists(app_icon_path):
            pix = QPixmap(app_icon_path).scaled(30, 30, Qt.KeepAspectRatio, Qt.SmoothTransformation)
            app_icon_label.setPixmap(pix)
        else:
            app_icon_label.setText("✦")
            app_icon_label.setFont(QFont("Segoe UI", 14, QFont.Bold))
            app_icon_label.setAlignment(Qt.AlignCenter)
            app_icon_label.setStyleSheet(f"color: {AppPalette.PRIMARY};")

        h_layout.addWidget(app_icon_label)

        title_col = QVBoxLayout()
        title_col.setSpacing(1)
        title_col.setContentsMargins(0, 0, 0, 0)

        app_title = QLabel("Copilots")
        app_title.setFont(QFont("Segoe UI", 12, QFont.Bold))
        app_title.setStyleSheet(f"color: {AppPalette.TEXT_PRIMARY};")
        title_col.addWidget(app_title)

        app_subtitle = QLabel("Unified Desktop Suite")
        app_subtitle.setFont(QFont("Segoe UI", 8))
        app_subtitle.setStyleSheet(f"color: {AppPalette.TEXT_MUTED};")
        title_col.addWidget(app_subtitle)

        h_layout.addLayout(title_col)
        h_layout.addStretch()
        layout.addWidget(header_frame)

        # Nav Items Area
        nav_container = QWidget()
        nav_layout = QVBoxLayout(nav_container)
        nav_layout.setContentsMargins(12, 16, 12, 16)
        nav_layout.setSpacing(6)

        section_title = QLabel("COPILOTS")
        section_title.setFont(QFont("Segoe UI", 8, QFont.Bold))
        section_title.setStyleSheet(f"color: {AppPalette.TEXT_MUTED}; padding-left: 6px; padding-bottom: 4px;")
        nav_layout.addWidget(section_title)

        self.nav_items: Dict[str, NavItem] = {
            "powerpoint": NavItem(
                route_id="powerpoint",
                title="PowerPoint Copilot",
                icon_path="icons/powerpoint.png",
                badge_color=AppPalette.BRAND_PPT,
                is_selected=(current_route == "powerpoint"),
            ),
            "word": NavItem(
                route_id="word",
                title="Word Copilot",
                icon_path="icons/word.png",
                badge_color=AppPalette.BRAND_WORD,
                is_selected=(current_route == "word"),
            ),
            "excel": NavItem(
                route_id="excel",
                title="Excel Copilot",
                icon_path="icons/excel.png",
                badge_color=AppPalette.BRAND_EXCEL,
                is_selected=(current_route == "excel"),
            ),
            "cv": NavItem(
                route_id="cv",
                title="CV Copilot",
                icon_path="icons/cv.png",
                badge_color=AppPalette.BRAND_CV,
                is_selected=(current_route == "cv"),
            ),
        }

        for k, item in self.nav_items.items():
            item.clicked.connect(self._on_item_clicked)
            nav_layout.addWidget(item)

        nav_layout.addStretch()
        layout.addWidget(nav_container)

        # Footer Area
        footer_frame = QFrame()
        footer_frame.setStyleSheet(f"""
            border-top: 1px solid {AppPalette.BORDER_COLOR};
        """)
        f_layout = QVBoxLayout(footer_frame)
        f_layout.setContentsMargins(12, 12, 12, 12)
        f_layout.setSpacing(6)

        manage_btn = QPushButton("Manage Prompts Folder")
        manage_btn.setFont(QFont("Segoe UI", 9))
        manage_btn.setCursor(Qt.PointingHandCursor)
        manage_btn.setStyleSheet(f"""
            QPushButton {{
                background-color: transparent;
                color: {AppPalette.TEXT_MUTED};
                border: none;
                padding: 4px;
                text-align: center;
            }}
            QPushButton:hover {{
                color: {AppPalette.TEXT_PRIMARY};
            }}
        """)
        manage_btn.clicked.connect(lambda: PromptManager().open_prompts_directory())
        f_layout.addWidget(manage_btn)

        local_badge = QLabel("🛡 Air-Gapped & Local")
        local_badge.setFont(QFont("Segoe UI", 8))
        local_badge.setAlignment(Qt.AlignCenter)
        local_badge.setStyleSheet(f"color: {AppPalette.SUCCESS};")
        f_layout.addWidget(local_badge)

        layout.addWidget(footer_frame)

    def _on_item_clicked(self, route_id: str):
        self.set_active_route(route_id)
        if self.on_navigate:
            self.on_navigate(route_id)

    def set_active_route(self, route_id: str):
        self.current_route = route_id
        for k, item in self.nav_items.items():
            item.set_selected(k == route_id)
