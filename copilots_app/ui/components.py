"""
Reusable UI Components: App Header, Status Bar, Code/DSL Editor, Metric Cards, Action Bars, and Dialogs.
"""

import flet as ft
from typing import Optional, Callable, List, Dict, Any
from copilots_app.core.theme import AppPalette


class AppHeader(ft.Container):
    """Modern header banner for Copilot views with title, subtitle, and badge."""

    def __init__(
        self,
        title: str,
        subtitle: str,
        icon_path: Optional[str] = None,
        icon_name: Optional[str] = None,
        badge_text: Optional[str] = None,
        badge_color: str = AppPalette.PRIMARY,
        actions: Optional[List[ft.Control]] = None,
    ):
        lead_controls = []
        if icon_path:
            lead_controls.append(
                ft.Container(
                    content=ft.Image(src=icon_path, width=32, height=32, fit=ft.BoxFit.COVER, border_radius=16),
                    padding=6,
                    border_radius=22,
                    bgcolor=AppPalette.BG_CARD,
                    border=ft.Border.all(1, AppPalette.BORDER_COLOR),
                )
            )
        elif icon_name:
            lead_controls.append(
                ft.Container(
                    content=ft.Icon(icon_name, size=24, color=AppPalette.TEXT_PRIMARY),
                    padding=6,
                    border_radius=22,
                    bgcolor=AppPalette.BG_CARD,
                    border=ft.Border.all(1, AppPalette.BORDER_COLOR),
                )
            )

        title_row = [
            ft.Text(title, size=18, weight=ft.FontWeight.BOLD, color=AppPalette.TEXT_PRIMARY),
        ]
        if badge_text:
            title_row.append(
                ft.Container(
                    content=ft.Text(badge_text, size=10, weight=ft.FontWeight.W_600, color="#FFFFFF"),
                    padding=ft.Padding.symmetric(horizontal=8, vertical=2),
                    border_radius=12,
                    bgcolor=badge_color,
                )
            )

        text_col = ft.Column(
            controls=[
                ft.Row(controls=title_row, spacing=8, vertical_alignment=ft.CrossAxisAlignment.CENTER),
                ft.Text(subtitle, size=12, color=AppPalette.TEXT_SECONDARY),
            ],
            spacing=2,
            alignment=ft.MainAxisAlignment.CENTER,
        )
        lead_controls.append(text_col)

        row_controls = [
            ft.Row(controls=lead_controls, spacing=12, vertical_alignment=ft.CrossAxisAlignment.CENTER)
        ]

        if actions:
            row_controls.append(ft.Row(controls=actions, spacing=8))

        super().__init__(
            content=ft.Row(
                controls=row_controls,
                alignment=ft.MainAxisAlignment.SPACE_BETWEEN,
                vertical_alignment=ft.CrossAxisAlignment.CENTER,
            ),
            padding=ft.Padding.symmetric(horizontal=16, vertical=12),
            bgcolor=AppPalette.BG_SURFACE,
            border=ft.Border.only(bottom=ft.BorderSide(1, AppPalette.BORDER_COLOR)),
        )


class StatusBar(ft.Container):
    """Async status indicator bar with spinner, levels and icons."""

    def __init__(self, default_text: str = "Ready"):
        self.icon = ft.Icon(ft.Icons.INFO_OUTLINE, size=16, color=AppPalette.TEXT_SECONDARY)
        self.text = ft.Text(default_text, size=12, color=AppPalette.TEXT_SECONDARY, expand=True)
        self.spinner = ft.ProgressRing(width=14, height=14, stroke_width=2, color=AppPalette.PRIMARY, visible=False)

        super().__init__(
            content=ft.Row(
                controls=[self.icon, self.spinner, self.text],
                spacing=8,
                vertical_alignment=ft.CrossAxisAlignment.CENTER,
            ),
            padding=ft.Padding.symmetric(horizontal=16, vertical=8),
            bgcolor=AppPalette.BG_SURFACE,
            border=ft.Border.only(top=ft.BorderSide(1, AppPalette.BORDER_COLOR)),
        )

    def set_status(self, message: str, level: str = "info", loading: bool = False):
        level_map = {
            "info": (ft.Icons.INFO_OUTLINE, AppPalette.TEXT_SECONDARY),
            "success": (ft.Icons.CHECK_CIRCLE_OUTLINE, AppPalette.SUCCESS),
            "warning": (ft.Icons.WARNING_AMBER_ROUNDED, AppPalette.WARNING),
            "error": (ft.Icons.ERROR_OUTLINE, AppPalette.ERROR),
        }
        icon_name, color = level_map.get(level, (ft.Icons.INFO_OUTLINE, AppPalette.TEXT_SECONDARY))

        self.spinner.visible = loading
        self.icon.visible = not loading
        self.icon.name = icon_name
        self.icon.color = color
        self.text.value = message
        self.text.color = color if level != "info" else AppPalette.TEXT_SECONDARY
        try:
            self.update()
        except Exception:
            pass


class CodeEditor(ft.Container):
    """Multi-line monospace code and DSL editor with toolbar."""

    def __init__(
        self,
        value: str = "",
        hint_text: str = "Paste or enter DSL code here...",
        on_change: Optional[Callable[[Any], None]] = None,
        samples: Optional[Dict[str, str]] = None,
        on_sample_selected: Optional[Callable[[str], None]] = None,
        expand: bool = True,
    ):
        self.text_field = ft.TextField(
            value=value,
            multiline=True,
            min_lines=15,
            max_lines=1000,
            text_size=13,
            text_style=ft.TextStyle(font_family="Consolas, Fira Code, monospace"),
            border=ft.InputBorder.NONE,
            filled=True,
            fill_color=AppPalette.BG_INPUT,
            cursor_color=AppPalette.PRIMARY,
            hint_text=hint_text,
            hint_style=ft.TextStyle(color=AppPalette.TEXT_MUTED, size=12),
            on_change=on_change,
            expand=True,
        )

        toolbar_items: List[ft.Control] = [
            ft.Text("DSL / Code Editor", size=11, weight=ft.FontWeight.W_600, color=AppPalette.TEXT_MUTED),
        ]

        if samples:
            self.sample_dropdown = ft.Dropdown(
                options=[ft.DropdownOption(k) for k in samples.keys()],
                hint_text="Load sample template...",
                text_size=11,
                content_padding=ft.Padding.symmetric(horizontal=8, vertical=4),
                border_color=AppPalette.BORDER_COLOR,
                bgcolor=AppPalette.BG_CARD,
                width=200,
                dense=True,
                on_select=lambda e: self._handle_sample_select(e, samples, on_sample_selected),
            )
            toolbar_items.append(self.sample_dropdown)

        toolbar_items.append(
            ft.IconButton(
                icon=ft.Icons.COPY_ALL,
                tooltip="Copy All to Clipboard",
                icon_size=16,
                icon_color=AppPalette.TEXT_MUTED,
                on_click=self._copy_all,
            )
        )
        toolbar_items.append(
            ft.IconButton(
                icon=ft.Icons.DELETE_SWEEP_OUTLINED,
                tooltip="Clear Editor",
                icon_size=16,
                icon_color=AppPalette.TEXT_MUTED,
                on_click=self._clear_all,
            )
        )

        editor_content = ft.Column(
            controls=[
                ft.Container(
                    content=ft.Row(
                        controls=[
                            toolbar_items[0],
                            ft.Row(controls=toolbar_items[1:], spacing=4, vertical_alignment=ft.CrossAxisAlignment.CENTER),
                        ],
                        alignment=ft.MainAxisAlignment.SPACE_BETWEEN,
                    ),
                    padding=ft.Padding.symmetric(horizontal=12, vertical=4),
                    bgcolor=AppPalette.BG_SURFACE,
                    border=ft.Border.only(bottom=ft.BorderSide(1, AppPalette.BORDER_COLOR)),
                ),
                self.text_field,
            ],
            spacing=0,
            expand=True,
        )

        super().__init__(
            content=editor_content,
            bgcolor=AppPalette.BG_INPUT,
            border=ft.Border.all(1, AppPalette.BORDER_COLOR),
            border_radius=8,
            clip_behavior=ft.ClipBehavior.ANTI_ALIAS,
            expand=expand,
        )

    def _handle_sample_select(self, e, samples: Dict[str, str], callback: Optional[Callable[[str], None]]):
        selected_key = e.control.value
        if selected_key in samples:
            self.set_value(samples[selected_key])
            if callback:
                callback(selected_key)

    def _copy_all(self, e):
        e.page.set_clipboard(self.text_field.value)
        e.page.show_snack_bar(ft.SnackBar(content=ft.Text("Copied editor text to clipboard"), bgcolor=AppPalette.SUCCESS))

    def _clear_all(self, e):
        self.set_value("")

    def get_value(self) -> str:
        return self.text_field.value or ""

    def set_value(self, val: str):
        self.text_field.value = val
        try:
            self.text_field.update()
        except Exception:
            pass


class MetricCard(ft.Container):
    """Sleek KPI / Summary Metric Card with accent icon."""

    def __init__(self, title: str, value: str, icon_name: str, color: str = AppPalette.PRIMARY, subtitle: str = ""):
        super().__init__(
            content=ft.Row(
                controls=[
                    ft.Container(
                        content=ft.Icon(icon_name, size=24, color=color),
                        padding=10,
                        border_radius=8,
                        bgcolor=AppPalette.BG_INPUT,
                        border=ft.Border.all(1, AppPalette.BORDER_COLOR),
                    ),
                    ft.Column(
                        controls=[
                            ft.Text(title, size=11, color=AppPalette.TEXT_SECONDARY),
                            ft.Text(value, size=18, weight=ft.FontWeight.BOLD, color=AppPalette.TEXT_PRIMARY),
                            ft.Text(subtitle, size=10, color=AppPalette.TEXT_MUTED) if subtitle else ft.Container(),
                        ],
                        spacing=1,
                    ),
                ],
                spacing=12,
                vertical_alignment=ft.CrossAxisAlignment.CENTER,
            ),
            padding=12,
            bgcolor=AppPalette.BG_CARD,
            border=ft.Border.all(1, AppPalette.BORDER_COLOR),
            border_radius=8,
            expand=True,
        )


class ActionButton(ft.Button):
    """Styled action button with loading support."""

    def __init__(
        self,
        text: str,
        icon: Optional[str] = None,
        color: str = AppPalette.PRIMARY,
        on_click: Optional[Callable[[Any], None]] = None,
        tooltip: Optional[str] = None,
        disabled: bool = False,
    ):
        super().__init__(
            content=text,
            icon=icon,
            color="#FFFFFF",
            bgcolor=color,
            on_click=on_click,
            tooltip=tooltip,
            disabled=disabled,
        )

