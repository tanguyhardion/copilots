"""
PowerPoint Copilot View: DSL Editor, Quick Templates, Presentation Controls, and Cheatsheet.
"""

import threading
import flet as ft
from typing import Optional

from copilots_app.core.theme import AppPalette
from copilots_app.ui.components import AppHeader, StatusBar, CodeEditor, ActionButton, MetricCard
from copilots_app.services.powerpoint import (
    PPT_SAMPLES,
    PowerPointConnector,
    parse_dsl_slides,
    refresh_dsl_theme_colors,
)


class PowerPointView(ft.Container):
    """Unified PowerPoint Copilot interface in Flet."""

    def __init__(self, page: ft.Page):
        self.app_page = page
        self.connector = PowerPointConnector()

        self.header = AppHeader(
            title="PowerPoint Copilot",
            subtitle="Generate slides, shapes, rich text, tables, and icons from clean DSL",
            icon_path="icons/powerpoint.png",
            badge_text="PowerPoint COM",
            badge_color=AppPalette.BRAND_PPT,
            actions=[
                ft.TextButton(
                    "DSL Cheatsheet",
                    icon=ft.Icons.HELP_OUTLINE,
                    on_click=self._show_cheatsheet,
                ),
            ],
        )

        self.status_bar = StatusBar(default_text="Ready — write or paste DSL and choose an action")

        self.editor = CodeEditor(
            value=PPT_SAMPLES["Overview Demo"],
            hint_text="Enter PowerPoint DSL here (rect, rounded_rect, table, icon, chevron...)",
            samples=PPT_SAMPLES,
            on_sample_selected=self._on_sample_loaded,
        )

        # Action Buttons
        self.copy_btn = ActionButton(
            text="Copy to Clipboard",
            icon=ft.Icons.CONTENT_COPY,
            color=AppPalette.BRAND_PPT,
            tooltip="Creates shapes in background and copies them to clipboard for Ctrl+V in PowerPoint",
            on_click=self._on_copy_clipboard,
        )

        self.insert_btn = ActionButton(
            text="Insert on Current Slide",
            icon=ft.Icons.ADD_TO_PHOTOS_OUTLINED,
            color=AppPalette.SUCCESS,
            tooltip="Directly injects shapes onto the currently selected slide in PowerPoint",
            on_click=self._on_insert_current_slide,
        )

        self.slide_btn = ActionButton(
            text="Create Full Slide(s)",
            icon=ft.Icons.NOTE_ADD_OUTLINED,
            color=AppPalette.PRIMARY,
            tooltip="Appends one or more new slides in the active PowerPoint presentation",
            on_click=self._on_create_full_slide,
        )

        action_row = ft.Container(
            content=ft.Row(
                controls=[
                    self.copy_btn,
                    self.insert_btn,
                    self.slide_btn,
                ],
                spacing=10,
                wrap=True,
            ),
            padding=ft.Padding.symmetric(horizontal=16, vertical=10),
            bgcolor=AppPalette.BG_SURFACE,
            border=ft.Border.only(top=ft.BorderSide(1, AppPalette.BORDER_COLOR)),
        )

        main_layout = ft.Column(
            controls=[
                self.header,
                ft.Container(
                    content=self.editor,
                    padding=16,
                    expand=True,
                ),
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

    def _set_buttons_state(self, disabled: bool):
        self.copy_btn.disabled = disabled
        self.insert_btn.disabled = disabled
        self.slide_btn.disabled = disabled
        try:
            self.copy_btn.update()
            self.insert_btn.update()
            self.slide_btn.update()
        except Exception:
            pass

    def _on_sample_loaded(self, sample_name: str):
        self.status_bar.set_status(f"Loaded sample template: '{sample_name}'", level="info")

    def _get_slides(self):
        text = self.editor.get_value().strip()
        if not text:
            return [[]]
        try:
            refresh_dsl_theme_colors()
        except Exception:
            pass
        return parse_dsl_slides(text)

    def _on_copy_clipboard(self, e):
        slides = self._get_slides()
        shapes = slides[0] if slides else []
        if not shapes:
            self.status_bar.set_status("DSL is empty — nothing to copy", level="warning")
            return

        self._set_buttons_state(True)
        self.status_bar.set_status("Building shapes for clipboard…", level="info", loading=True)

        def worker():
            try:
                self.connector.create_shapes_and_copy(
                    shapes,
                    status_cb=lambda m: self.status_bar.set_status(m, level="info", loading=True),
                )
                self.status_bar.set_status("✓ Shapes copied to clipboard — switch to PowerPoint and press Ctrl+V", level="success")
            except Exception as err:
                self.status_bar.set_status(f"Clipboard copy failed: {err}", level="error")
            finally:
                self._set_buttons_state(False)

        threading.Thread(target=worker, daemon=True).start()

    def _on_insert_current_slide(self, e):
        slides = self._get_slides()
        shapes = slides[0] if slides else []
        if not shapes:
            self.status_bar.set_status("DSL is empty — nothing to insert", level="warning")
            return

        self._set_buttons_state(True)
        self.status_bar.set_status("Inserting shapes onto active slide…", level="info", loading=True)

        def worker():
            try:
                self.connector.create_on_current_slide(
                    shapes,
                    status_cb=lambda m: self.status_bar.set_status(m, level="info", loading=True),
                )
                self.status_bar.set_status("✓ Shapes successfully inserted on active slide!", level="success")
            except Exception as err:
                self.status_bar.set_status(f"Insert failed: {err}", level="error")
            finally:
                self._set_buttons_state(False)

        threading.Thread(target=worker, daemon=True).start()

    def _on_create_full_slide(self, e):
        slides = self._get_slides()
        total_shapes = sum(len(s) for s in slides)
        if total_shapes == 0:
            self.status_bar.set_status("DSL is empty — nothing to build", level="warning")
            return

        self._set_buttons_state(True)
        self.status_bar.set_status("Creating new slide(s)…", level="info", loading=True)

        def worker():
            try:
                self.connector.create_on_new_slide(
                    slides,
                    status_cb=lambda m: self.status_bar.set_status(m, level="info", loading=True),
                )
                self.status_bar.set_status(f"✓ Created {len(slides)} slide(s) ({total_shapes} shapes total) successfully!", level="success")
            except Exception as err:
                self.status_bar.set_status(f"Slide creation failed: {err}", level="error")
            finally:
                self._set_buttons_state(False)

        threading.Thread(target=worker, daemon=True).start()

    def _show_cheatsheet(self, e):
        dialog = ft.AlertDialog(
            title=ft.Text("PowerPoint DSL Cheatsheet", size=16, weight=ft.FontWeight.BOLD),
            content=ft.Container(
                content=ft.Column(
                    controls=[
                        ft.Text("Coordinate System: 960 × 540pt (16:9 Standard Slide)", weight=ft.FontWeight.BOLD),
                        ft.Text("• Margins: left=48, right=912, top=48, bottom=492"),
                        ft.Divider(color=AppPalette.BORDER_COLOR),
                        ft.Text("Syntax Examples:", weight=ft.FontWeight.BOLD),
                        ft.Text("rect left=40 top=40 width=200 height=100 color=a1 | \"Card Title\" size=14 bold=true"),
                        ft.Text("rounded_rect left=260 top=40 width=200 height=100 border_radius=10 color=a6 outline=a1,2"),
                        ft.Text("icon name=chart-line style=solid left=60 top=60 width=32 height=32 color=a1"),
                        ft.Text("table left=48 top=200 width=864 height=200 cols=200,400,264 header=\"A\",\"B\",\"C\" row=\"1\",\"2\",\"3\""),
                        ft.Text("line x1=48 y1=450 x2=912 y2=450 color=a3 weight=1.5 dash=solid"),
                        ft.Divider(color=AppPalette.BORDER_COLOR),
                        ft.Text("Slide Separator: Use '---' on a separate line for multi-slide generation."),
                    ],
                    spacing=8,
                    scroll=ft.ScrollMode.AUTO,
                ),
                width=540,
                height=380,
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

