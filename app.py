"""
Main PySide6 Application Entrypoint for the Unified Copilot Suite.
"""

import os
import sys
from PySide6.QtWidgets import (
    QApplication,
    QMainWindow,
    QWidget,
    QHBoxLayout,
    QStackedWidget,
)
from PySide6.QtGui import QIcon
from PySide6.QtCore import Qt

# Add current workspace to path
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from copilots_app.core.theme import get_app_stylesheet, get_asset_path
from copilots_app.ui.sidebar import AppSidebar
from copilots_app.ui.views.powerpoint_view import PowerPointView
from copilots_app.ui.views.word_view import WordView
from copilots_app.ui.views.excel_view import ExcelView
from copilots_app.ui.views.cv_view import CVCopilotView


class MainWindow(QMainWindow):
    """Main Application Window housing the sidebar and stacked copilot views."""

    def __init__(self):
        super().__init__()
        self.setWindowTitle("Copilots")
        self.setMinimumSize(1000, 680)

        # Set Window App Icon
        icon_path = get_asset_path("icons/copilots.png")
        if os.path.exists(icon_path):
            self.setWindowIcon(QIcon(icon_path))

        # Central Widget & Layout
        central_widget = QWidget()
        self.setCentralWidget(central_widget)

        root_layout = QHBoxLayout(central_widget)
        root_layout.setContentsMargins(0, 0, 0, 0)
        root_layout.setSpacing(0)

        # Content Area Stack
        self.stack = QStackedWidget()

        self.views = {
            "powerpoint": PowerPointView(self),
            "word": WordView(self),
            "excel": ExcelView(self),
            "cv": CVCopilotView(self),
        }

        self.view_indices = {}
        for key, view_widget in self.views.items():
            idx = self.stack.addWidget(view_widget)
            self.view_indices[key] = idx

        # Sidebar
        self.sidebar = AppSidebar(
            current_route="powerpoint",
            on_navigate=self.navigate_to,
            parent=self,
        )

        root_layout.addWidget(self.sidebar)
        root_layout.addWidget(self.stack, 1)

        # Set initial view
        self.navigate_to("powerpoint")

    def navigate_to(self, route_id: str):
        if route_id in self.view_indices:
            self.stack.setCurrentIndex(self.view_indices[route_id])
            self.sidebar.set_active_route(route_id)


def main():
    app = QApplication(sys.argv)
    app.setStyle("Fusion")
    app.setStyleSheet(get_app_stylesheet())

    window = MainWindow()
    window.showMaximized()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
