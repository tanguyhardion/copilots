"""
Sidebar navigation menu with brand icons, active states, and quick settings toggle.
"""

import flet as ft
from typing import Callable, Optional
from copilots_app.core.theme import AppPalette


class NavItem(ft.Container):
    """Custom navigation sidebar item with icon badge and hover effects."""

    def __init__(
        self,
        route_id: str,
        title: str,
        icon_path: Optional[str] = None,
        icon_name: Optional[str] = None,
        badge_color: str = AppPalette.PRIMARY,
        is_selected: bool = False,
        on_select: Optional[Callable[[str], None]] = None,
    ):
        self.route_id = route_id
        self.on_select = on_select
        self.is_selected = is_selected
        self.badge_color = badge_color

        icon_widget = (
            ft.Image(src=icon_path, width=22, height=22, fit=ft.BoxFit.CONTAIN)
            if icon_path
            else ft.Icon(icon_name or ft.Icons.APPS, size=20, color=AppPalette.TEXT_PRIMARY)
        )

        self.label = ft.Text(
            title,
            size=13,
            weight=ft.FontWeight.W_600 if is_selected else ft.FontWeight.NORMAL,
            color=AppPalette.TEXT_PRIMARY if is_selected else AppPalette.TEXT_SECONDARY,
        )

        self.indicator = ft.Container(
            width=3,
            height=20,
            border_radius=2,
            bgcolor=badge_color if is_selected else "transparent",
        )

        content = ft.Row(
            controls=[
                self.indicator,
                icon_widget,
                self.label,
            ],
            spacing=10,
            vertical_alignment=ft.CrossAxisAlignment.CENTER,
        )

        super().__init__(
            content=content,
            padding=ft.Padding.symmetric(horizontal=8, vertical=10),
            border_radius=6,
            bgcolor=AppPalette.BG_CARD if is_selected else "transparent",
            on_click=self._handle_click,
            on_hover=self._handle_hover,
            animate=ft.Animation(150, ft.AnimationCurve.EASE_OUT),
        )

    def _handle_click(self, e):
        if self.on_select:
            self.on_select(self.route_id)

    def _handle_hover(self, e):
        if not self.is_selected:
            self.bgcolor = AppPalette.BG_CARD_HOVER if e.data == "true" else "transparent"
            self.update()

    def set_selected(self, selected: bool):
        self.is_selected = selected
        self.bgcolor = AppPalette.BG_CARD if selected else "transparent"
        self.indicator.bgcolor = self.badge_color if selected else "transparent"
        self.label.weight = ft.FontWeight.W_600 if selected else ft.FontWeight.NORMAL
        self.label.color = AppPalette.TEXT_PRIMARY if selected else AppPalette.TEXT_SECONDARY
        try:
            self.update()
        except Exception:
            pass


class AppSidebar(ft.Container):
    """Main application sidebar containing brand headers, navigation items, and settings."""

    def __init__(self, current_route: str = "powerpoint", on_navigate: Optional[Callable[[str], None]] = None):
        self.on_navigate = on_navigate
        self.current_route = current_route

        self.nav_items = {
            "powerpoint": NavItem(
                route_id="powerpoint",
                title="PowerPoint Copilot",
                icon_path="icons/powerpoint.png",
                badge_color=AppPalette.BRAND_PPT,
                is_selected=(current_route == "powerpoint"),
                on_select=self._on_item_selected,
            ),
            "word": NavItem(
                route_id="word",
                title="Word Copilot",
                icon_path="icons/word.png",
                badge_color=AppPalette.BRAND_WORD,
                is_selected=(current_route == "word"),
                on_select=self._on_item_selected,
            ),
            "excel": NavItem(
                route_id="excel",
                title="Excel Copilot",
                icon_path="icons/excel.png",
                badge_color=AppPalette.BRAND_EXCEL,
                is_selected=(current_route == "excel"),
                on_select=self._on_item_selected,
            ),
            "cv": NavItem(
                route_id="cv",
                title="CV Copilot",
                icon_path="icons/cv.png",
                badge_color=AppPalette.BRAND_CV,
                is_selected=(current_route == "cv"),
                on_select=self._on_item_selected,
            ),
        }

        # App Brand Header
        brand_header = ft.Container(
            content=ft.Row(
                controls=[
                    ft.Container(
                        content=ft.Icon(ft.Icons.AUTO_AWESOME, size=20, color=AppPalette.PRIMARY),
                        padding=8,
                        border_radius=8,
                        bgcolor=AppPalette.BG_CARD,
                        border=ft.Border.all(1, AppPalette.BORDER_COLOR),
                    ),
                    ft.Column(
                        controls=[
                            ft.Text("Copilots", size=15, weight=ft.FontWeight.BOLD, color=AppPalette.TEXT_PRIMARY),
                            ft.Text("Unified Desktop Suite", size=10, color=AppPalette.TEXT_MUTED),
                        ],
                        spacing=0,
                    ),
                ],
                spacing=10,
                vertical_alignment=ft.CrossAxisAlignment.CENTER,
            ),
            padding=ft.Padding.symmetric(horizontal=12, vertical=16),
            border=ft.Border.only(bottom=ft.BorderSide(1, AppPalette.BORDER_COLOR)),
        )

        nav_column = ft.Column(
            controls=[
                ft.Container(
                    content=ft.Text("COPILOTS", size=10, weight=ft.FontWeight.BOLD, color=AppPalette.TEXT_MUTED),
                    padding=ft.Padding.only(left=12, top=12, bottom=6),
                ),
                self.nav_items["powerpoint"],
                self.nav_items["word"],
                self.nav_items["excel"],
                self.nav_items["cv"],
            ],
            spacing=4,
            expand=True,
        )

        footer = ft.Container(
            content=ft.Row(
                controls=[
                    ft.Icon(ft.Icons.SHIELD_OUTLINED, size=14, color=AppPalette.SUCCESS),
                    ft.Text("Air-Gapped & Local", size=11, color=AppPalette.TEXT_MUTED),
                ],
                spacing=6,
                alignment=ft.MainAxisAlignment.CENTER,
            ),
            padding=ft.Padding.symmetric(vertical=12),
            border=ft.Border.only(top=ft.BorderSide(1, AppPalette.BORDER_COLOR)),
        )

        sidebar_content = ft.Column(
            controls=[brand_header, nav_column, footer],
            spacing=0,
            expand=True,
        )

        super().__init__(
            content=sidebar_content,
            width=240,
            bgcolor=AppPalette.BG_SURFACE,
            border=ft.Border.only(right=ft.BorderSide(1, AppPalette.BORDER_COLOR)),
        )

    def _on_item_selected(self, route_id: str):
        self.set_active_route(route_id)
        if self.on_navigate:
            self.on_navigate(route_id)

    def set_active_route(self, route_id: str):
        self.current_route = route_id
        for k, item in self.nav_items.items():
            item.set_selected(k == route_id)
