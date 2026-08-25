"""
Prompt Management Dialog: View, Edit, Copy, Save Custom, and Reset System Prompts.
"""

import flet as ft
from typing import Callable, Optional
from copilots_app.core.theme import AppPalette
from copilots_app.core.prompt_manager import PromptManager


class PromptDialog(ft.AlertDialog):
    """Modern modal dialog for editing, copying, and resetting LLM System Prompts."""

    def __init__(
        self,
        page: ft.Page,
        copilot_key: str,
        on_status_change: Optional[Callable[[str, str], None]] = None,
    ):
        self.app_page = page
        self.copilot_key = copilot_key
        self.on_status_change = on_status_change
        self.pm = PromptManager()

        title_text = self.pm.PROMPT_TITLES.get(copilot_key, f"{copilot_key.capitalize()} System Prompt")
        self.is_custom = self.pm.is_customized(copilot_key)

        # Status badge for customized vs default
        self.badge_text = ft.Text(
            "User Custom Override" if self.is_custom else "Default Bundled",
            size=11,
            weight=ft.FontWeight.W_600,
            color="#FFFFFF" if self.is_custom else AppPalette.TEXT_SECONDARY,
        )
        self.badge_container = ft.Container(
            content=self.badge_text,
            padding=ft.Padding.symmetric(horizontal=8, vertical=3),
            border_radius=12,
            bgcolor=AppPalette.WARNING if self.is_custom else AppPalette.BG_SURFACE,
            border=ft.Border.all(1, AppPalette.BORDER_COLOR),
        )

        current_prompt = self.pm.get_prompt(copilot_key)

        # Text editor
        self.editor = ft.TextField(
            value=current_prompt,
            multiline=True,
            min_lines=18,
            max_lines=24,
            text_size=12,
            text_style=ft.TextStyle(font_family="Consolas, 'Cascadia Code', monospace"),
            border=ft.InputBorder.OUTLINE,
            border_color=AppPalette.BORDER_COLOR,
            filled=True,
            fill_color=AppPalette.BG_INPUT,
            cursor_color=AppPalette.PRIMARY,
            expand=True,
        )

        # Info row
        info_row = ft.Row(
            controls=[
                ft.Row(
                    controls=[
                        ft.Icon(ft.Icons.PSYCHOLOGY, size=18, color=AppPalette.PRIMARY),
                        ft.Text(title_text, size=15, weight=ft.FontWeight.BOLD, color=AppPalette.TEXT_PRIMARY),
                        self.badge_container,
                    ],
                    spacing=8,
                    vertical_alignment=ft.CrossAxisAlignment.CENTER,
                ),
                ft.TextButton(
                    "Open Prompts Folder",
                    icon=ft.Icons.FOLDER_OPEN_OUTLINED,
                    on_click=lambda _: self.pm.open_prompts_directory(),
                ),
            ],
            alignment=ft.MainAxisAlignment.SPACE_BETWEEN,
        )

        # Actions row with clear horizontal spacing
        actions_bar = ft.Container(
            content=ft.Row(
                controls=[
                    ft.TextButton(
                        "Reset to Default",
                        icon=ft.Icons.RESTART_ALT,
                        tooltip="Revert custom prompt back to factory bundled default",
                        on_click=self._on_reset,
                    ),
                    ft.Row(
                        controls=[
                            ft.OutlinedButton(
                                "Copy to Clipboard",
                                icon=ft.Icons.COPY_ALL,
                                on_click=self._on_copy,
                            ),
                            ft.ElevatedButton(
                                "Save Changes",
                                icon=ft.Icons.SAVE,
                                bgcolor=AppPalette.PRIMARY,
                                color="#FFFFFF",
                                on_click=self._on_save,
                            ),
                            ft.TextButton(
                                "Close",
                                on_click=self._on_close,
                            ),
                        ],
                        spacing=10,
                    ),
                ],
                alignment=ft.MainAxisAlignment.SPACE_BETWEEN,
                vertical_alignment=ft.CrossAxisAlignment.CENTER,
            ),
            padding=ft.Padding.only(top=10),
            expand=True,
        )

        super().__init__(
            title=info_row,
            content=ft.Container(
                content=ft.Column(
                    controls=[
                        ft.Text(
                            "This system prompt instructs the LLM how to format responses and DSL code for this copilot.",
                            size=12,
                            color=AppPalette.TEXT_MUTED,
                        ),
                        self.editor,
                    ],
                    spacing=10,
                ),
                width=850,
                height=530,
            ),
            actions=[actions_bar],
            modal=True,
        )

    def _on_copy(self, e):
        text = self.editor.value or ""
        self.app_page.set_clipboard(text)
        if self.on_status_change:
            self.on_status_change("System prompt copied to clipboard!", "info")
        
        # Show a snackbar or quick notice
        snack = ft.SnackBar(ft.Text("✓ System prompt copied to clipboard!"), open=True)
        self.app_page.overlay.append(snack)
        self.app_page.update()

    def _on_save(self, e):
        content = self.editor.value or ""
        ok = self.pm.save_user_prompt(self.copilot_key, content)
        if ok:
            self.badge_text.value = "User Custom Override"
            self.badge_text.color = "#FFFFFF"
            self.badge_container.bgcolor = AppPalette.WARNING
            self.badge_container.update()
            if self.on_status_change:
                self.on_status_change("Saved custom system prompt override in AppData", "success")
            snack = ft.SnackBar(ft.Text("✓ Custom prompt saved to AppData!"), open=True)
            self.app_page.overlay.append(snack)
            self.app_page.update()

    def _on_reset(self, e):
        self.pm.reset_to_default(self.copilot_key)
        default_text = self.pm.get_default_prompt(self.copilot_key)
        self.editor.value = default_text
        self.editor.update()

        self.badge_text.value = "Default Bundled"
        self.badge_text.color = AppPalette.TEXT_SECONDARY
        self.badge_container.bgcolor = AppPalette.BG_SURFACE
        self.badge_container.update()

        if self.on_status_change:
            self.on_status_change("Restored factory default prompt", "info")
        snack = ft.SnackBar(ft.Text("✓ Reset to factory default prompt!"), open=True)
        self.app_page.overlay.append(snack)
        self.app_page.update()

    def _on_close(self, e):
        self.open = False
        self.app_page.update()


def open_prompt_dialog(
    page: ft.Page,
    copilot_key: str,
    on_status_change: Optional[Callable[[str, str], None]] = None,
):
    """Helper to instantiate and show PromptDialog on a page."""
    dialog = PromptDialog(page, copilot_key, on_status_change)
    page.overlay.append(dialog)
    dialog.open = True
    page.update()
