"""
Main Flet Application Entrypoint for the Unified Copilot Suite.
"""

import os
import sys
import flet as ft

# Add current workspace to path
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from copilots_app.core.theme import create_app_theme, AppPalette
from copilots_app.ui.sidebar import AppSidebar
from copilots_app.ui.views.powerpoint_view import PowerPointView
from copilots_app.ui.views.word_view import WordView
from copilots_app.ui.views.excel_view import ExcelView
from copilots_app.ui.views.cv_view import CVCopilotView


def main(page: ft.Page):
    page.title = "Copilots"
    page.window.maximized = True
    page.window.min_width = 1000
    page.window.min_height = 680
    page.theme = create_app_theme()
    page.theme_mode = ft.ThemeMode.DARK
    page.bgcolor = AppPalette.BG_DARK
    page.padding = 0

    # Views cache
    views = {
        "powerpoint": PowerPointView(page),
        "word": WordView(page),
        "excel": ExcelView(page),
        "cv": CVCopilotView(page),
    }

    content_area = ft.Container(
        content=views["powerpoint"],
        expand=True,
    )

    def on_navigate(route_id: str):
        if route_id in views:
            content_area.content = views[route_id]
            content_area.update()

    sidebar = AppSidebar(current_route="powerpoint", on_navigate=on_navigate)

    root_layout = ft.Row(
        controls=[
            sidebar,
            content_area,
        ],
        spacing=0,
        expand=True,
    )

    page.add(root_layout)


if __name__ == "__main__":
    ft.run(main, assets_dir="assets")
