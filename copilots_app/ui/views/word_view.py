"""
Word Copilot View: Rich DSL & Edit-Mode Editor, Build & Open, Insert, Apply Edits, Extract.
"""

import threading
import flet as ft
from typing import Optional

from copilots_app.core.theme import AppPalette
from copilots_app.ui.components import AppHeader, StatusBar, CodeEditor, ActionButton
from copilots_app.ui.prompt_dialog import open_prompt_dialog
from copilots_app.services.word.connector import WordConnector
from copilots_app.services.word.editor import WordEditor
from copilots_app.services.word.extractor import WordExtractor
from copilots_app.services.word.colors import refresh_word_theme_colors
from copilots_app.services.word.dsl.parser import parse_dsl_pages
from copilots_app.services.word.dsl.edit_parser import parse_edit_dsl
from copilots_app.services.word.gui.sample import SAMPLE_DSL

WORD_SAMPLES = {
    "Complete Demo": SAMPLE_DSL,
    "Executive Report": """\
// Document setup
page size=a4 orientation=portrait margin=54,54,54,54

h1 align=center color=a4 | "Quarterly Strategy Review"
h3 align=center color=a2 | "Digital Operations & Modernization Program"

hr color=a3 weight=2

p spacing_after=8 | "This executive summary details the operational performance, technical deliverables, and roadmap milestones accomplished during the current cycle."

// Deliverables Table
table width=100% header_fill=a4 header_text_color=#FFFFFF text_color=t1 border_color=a3 border_weight=1 header_bold=true
cols=30%,45%,25%
header="Stream","Strategic Objective","Status"
row="Cloud Architecture","Scalable microservices migration","✓ Complete"
row="Automation Suite","Desktop Copilots unification in Flet","✓ Production"
row="Data Quality","Automated compliance validation","✓ Verified"

br

h2 color=a4 | "Key Recommendations"
ul indent=1
  item | "Maintain strict deterministic data validation across all output formats."
  item | "Accelerate user enablement sessions with interactive cheatsheets."
  item | "Continuous integration testing for COM office automation hooks."
""",
    "Edit Mode Example": """\
// Target active document in Microsoft Word
edit target=active

// Search and replace or update headings
replace find="Draft" replace="Final Approved Version"
insert_before find="Key Recommendations" | "Note: All milestones reviewed by Executive Sponsor."
""",
}


class WordView(ft.Container):
    """Unified Word Copilot interface in Flet."""

    def __init__(self, page: ft.Page):
        self.app_page = page
        self.connector = WordConnector()
        self.extractor = WordExtractor()
        self.editor = WordEditor()

        self.header = AppHeader(
            title="Word Copilot",
            subtitle="Build formatted documents, insert rich content at cursor, or live-edit active Word documents",
            icon_path="icons/word.png",
            badge_text="Word COM + python-docx",
            badge_color=AppPalette.BRAND_WORD,
            actions=[
                ft.TextButton(
                    "System Prompt",
                    icon=ft.Icons.PSYCHOLOGY_OUTLINED,
                    on_click=lambda _: open_prompt_dialog(
                        self.app_page,
                        "word",
                        on_status_change=lambda msg, lvl: self.status_bar.set_status(msg, level=lvl),
                    ),
                ),
                ft.TextButton(
                    "Word DSL Syntax",
                    icon=ft.Icons.HELP_OUTLINE,
                    on_click=self._show_syntax_help,
                ),
            ],
        )

        self.status_bar = StatusBar(default_text="Ready — write build DSL or edit-mode DSL")

        self.code_editor = CodeEditor(
            value=SAMPLE_DSL,
            hint_text="Paste Word build DSL (h1, p, table, ul, icon...) or edit DSL (edit target=active...)",
            samples=WORD_SAMPLES,
            on_sample_selected=self._on_sample_loaded,
        )

        # Action buttons
        self.open_btn = ActionButton(
            text="Build & Open in Word",
            icon=ft.Icons.FOLDER_OPEN_OUTLINED,
            color=AppPalette.BRAND_WORD,
            tooltip="Generates document from DSL and opens it in Microsoft Word",
            on_click=self._on_build_and_open,
        )

        self.insert_btn = ActionButton(
            text="Insert at Cursor",
            icon=ft.Icons.POST_ADD,
            color=AppPalette.SUCCESS,
            tooltip="Injects parsed elements directly at the current selection cursor in Word",
            on_click=self._on_insert_at_cursor,
        )

        self.apply_btn = ActionButton(
            text="Apply Edits to Word",
            icon=ft.Icons.EDIT_DOCUMENT,
            color="#5C3D8B",
            tooltip="Applies targeted edit operations to the active Word document",
            on_click=self._on_apply_edits,
        )

        self.extract_btn = ActionButton(
            text="Extract DSL from Word",
            icon=ft.Icons.DOCUMENT_SCANNER_OUTLINED,
            color=AppPalette.WARNING,
            tooltip="Reads the active Microsoft Word document and converts it to DSL in the editor",
            on_click=self._on_extract_dsl,
        )

        self.save_btn = ActionButton(
            text="Save .docx",
            icon=ft.Icons.SAVE_OUTLINED,
            color="#4A4F57",
            tooltip="Saves DSL directly to a .docx file",
            on_click=self._on_save_docx,
        )

        action_row = ft.Container(
            content=ft.Row(
                controls=[
                    self.open_btn,
                    self.insert_btn,
                    self.apply_btn,
                    self.extract_btn,
                    self.save_btn,
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
                    content=self.code_editor,
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
        for btn in (self.open_btn, self.insert_btn, self.apply_btn, self.extract_btn, self.save_btn):
            btn.disabled = disabled
            try:
                btn.update()
            except Exception:
                pass

    def _on_sample_loaded(self, sample_name: str):
        self.status_bar.set_status(f"Loaded Word template: '{sample_name}'", level="info")

    def _get_pages(self):
        text = self.code_editor.get_value().strip()
        if not text:
            return [[]]
        try:
            refresh_word_theme_colors()
        except Exception:
            pass
        return parse_dsl_pages(text)

    def _get_edit_ops(self):
        text = self.code_editor.get_value().strip()
        try:
            refresh_word_theme_colors()
        except Exception:
            pass
        return parse_edit_dsl(text)

    def _on_build_and_open(self, e):
        pages = self._get_pages()
        if not any(pages):
            self.status_bar.set_status("DSL is empty — nothing to build", level="warning")
            return

        self._set_buttons_state(True)
        self.status_bar.set_status("Building document and opening Word…", level="info", loading=True)

        def worker():
            try:
                self.connector.build_and_open(
                    pages,
                    status_cb=lambda m: self.status_bar.set_status(m, level="info", loading=True),
                )
                self.status_bar.set_status("✓ Document built and opened in Microsoft Word!", level="success")
            except Exception as err:
                self.status_bar.set_status(f"Build failed: {err}", level="error")
            finally:
                self._set_buttons_state(False)

        threading.Thread(target=worker, daemon=True).start()

    def _on_insert_at_cursor(self, e):
        pages = self._get_pages()
        if not any(pages):
            self.status_bar.set_status("DSL is empty — nothing to insert", level="warning")
            return

        self._set_buttons_state(True)
        self.status_bar.set_status("Inserting content at Word cursor…", level="info", loading=True)

        def worker():
            try:
                self.connector.insert_at_cursor(
                    pages,
                    status_cb=lambda m: self.status_bar.set_status(m, level="info", loading=True),
                )
                self.status_bar.set_status("✓ Content successfully inserted at cursor in Word!", level="success")
            except Exception as err:
                self.status_bar.set_status(f"Insert failed: {err}", level="error")
            finally:
                self._set_buttons_state(False)

        threading.Thread(target=worker, daemon=True).start()

    def _on_apply_edits(self, e):
        ops = self._get_edit_ops()
        if not ops:
            self.status_bar.set_status("No valid edit operations found in DSL", level="warning")
            return

        self._set_buttons_state(True)
        self.status_bar.set_status("Applying edits to active Word document…", level="info", loading=True)

        def worker():
            try:
                self.editor.apply_edits(
                    ops,
                    status_cb=lambda m: self.status_bar.set_status(m, level="info", loading=True),
                )
                self.status_bar.set_status("✓ Edit operations applied successfully!", level="success")
            except Exception as err:
                self.status_bar.set_status(f"Apply edits failed: {err}", level="error")
            finally:
                self._set_buttons_state(False)

        threading.Thread(target=worker, daemon=True).start()

    def _on_extract_dsl(self, e):
        self._set_buttons_state(True)
        self.status_bar.set_status("Extracting content from active Word document…", level="info", loading=True)

        def worker():
            try:
                dsl_text = self.extractor.extract(
                    status_cb=lambda m: self.status_bar.set_status(m, level="info", loading=True),
                )
                if dsl_text:
                    self.code_editor.set_value(dsl_text)
                    self.status_bar.set_status("✓ Successfully extracted DSL from Word!", level="success")
                else:
                    self.status_bar.set_status("No content extracted from active document", level="warning")
            except Exception as err:
                self.status_bar.set_status(f"Extraction failed: {err}", level="error")
            finally:
                self._set_buttons_state(False)

        threading.Thread(target=worker, daemon=True).start()

    def _on_save_docx(self, e):
        pages = self._get_pages()
        if not any(pages):
            self.status_bar.set_status("DSL is empty — nothing to save", level="warning")
            return

        self._set_buttons_state(True)
        self.status_bar.set_status("Saving document as .docx…", level="info", loading=True)

        def worker():
            try:
                out_file = "Generated_Word_Document.docx"
                self.connector.save_to_file(
                    pages,
                    out_file,
                    status_cb=lambda m: self.status_bar.set_status(m, level="info", loading=True),
                )
                self.status_bar.set_status(f"✓ Saved document to '{out_file}'", level="success")
            except Exception as err:
                self.status_bar.set_status(f"Save failed: {err}", level="error")
            finally:
                self._set_buttons_state(False)

        threading.Thread(target=worker, daemon=True).start()

    def _show_syntax_help(self, e):
        dialog = ft.AlertDialog(
            title=ft.Text("Word DSL Specification", size=16, weight=ft.FontWeight.BOLD),
            content=ft.Container(
                content=ft.Column(
                    controls=[
                        ft.Text("Structure Directives:", weight=ft.FontWeight.BOLD),
                        ft.Text("page size=letter|a4 orientation=portrait|landscape margin=72,72,72,72"),
                        ft.Text("h1..h6 align=left|center|right color=a1..a6 | \"Heading Text\""),
                        ft.Text("p spacing_after=6 | \"Text with \" + \"bold\" bold=true + \" parts\""),
                        ft.Divider(color=AppPalette.BORDER_COLOR),
                        ft.Text("Lists & Tables:", weight=ft.FontWeight.BOLD),
                        ft.Text("ul / ol indent=1 followed by indented item | \"bullet\""),
                        ft.Text("table width=100% header_fill=a4 text_color=t1 ... cols=... header=... row=..."),
                        ft.Divider(color=AppPalette.BORDER_COLOR),
                        ft.Text("Edit DSL (for active documents):", weight=ft.FontWeight.BOLD),
                        ft.Text("edit target=active"),
                        ft.Text("replace find=\"Old\" replace=\"New\""),
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

