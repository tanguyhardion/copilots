"""
Bridge API exposed to pywebview JavaScript frontend (window.pywebview.api).
Provides full access to all Copilots engines: PowerPoint, Word, Excel, CV, and System Prompts.
"""

import os
import json
import traceback
from typing import Dict, Any, List, Optional
import webview

from copilots_app.core.prompt_manager import PromptManager

# PowerPoint services
from copilots_app.services.powerpoint import (
    PowerPointConnector,
    parse_dsl_slides,
    refresh_dsl_theme_colors,
)

# Word services
from copilots_app.services.word.connector import WordConnector
from copilots_app.services.word.editor import WordEditor
from copilots_app.services.word.extractor import WordExtractor
from copilots_app.services.word.colors import refresh_word_theme_colors
from copilots_app.services.word.dsl.parser import parse_dsl_pages
from copilots_app.services.word.dsl.edit_parser import parse_edit_dsl

# Excel services
from copilots_app.services.excel.analyzer.workbook_analyzer import WorkbookAnalyzer
from copilots_app.services.excel.executor.executor_main import ActionExecutor
from copilots_app.services.excel.protocol.action_parser import ActionParser
from copilots_app.services.excel.models.protocol import ActionProtocol

# CV services
from copilots_app.services.cv import run_dq_audit, generate_cv

# Python Copilot services
from copilots_app.services.python_copilot.runner import PythonSandboxRunner


class CopilotBridge:
    """Python API bridge passed to pywebview."""

    def __init__(self):
        self._window = None
        self.prompt_manager = PromptManager()
        self.ppt_connector = PowerPointConnector()
        self.word_connector = WordConnector()
        self.word_extractor = WordExtractor()
        self.word_editor = WordEditor()
        self.excel_executor = ActionExecutor()
        self.excel_current_path: Optional[str] = None
        self.excel_current_model = None
        self.python_runner = PythonSandboxRunner()

    def set_window(self, window):
        self._window = window

    # -------------------------------------------------------------------------
    # System Prompts API
    # -------------------------------------------------------------------------
    def get_prompt_info(self, copilot_key: str) -> Dict[str, Any]:
        try:
            content = self.prompt_manager.get_prompt(copilot_key)
            is_custom = self.prompt_manager.is_customized(copilot_key)
            title = self.prompt_manager.PROMPT_TITLES.get(copilot_key, f"{copilot_key.capitalize()} Prompt")
            return {"success": True, "title": title, "content": content, "is_custom": is_custom}
        except Exception as e:
            return {"success": False, "error": str(e)}

    def save_prompt_override(self, copilot_key: str, content: str) -> Dict[str, Any]:
        try:
            self.prompt_manager.save_user_prompt(copilot_key, content)
            return {"success": True, "message": "Custom prompt override saved."}
        except Exception as e:
            return {"success": False, "error": str(e)}

    def reset_prompt_default(self, copilot_key: str) -> Dict[str, Any]:
        try:
            self.prompt_manager.reset_to_default(copilot_key)
            content = self.prompt_manager.get_prompt(copilot_key)
            return {"success": True, "content": content, "message": "Prompt reset to default."}
        except Exception as e:
            return {"success": False, "error": str(e)}

    # -------------------------------------------------------------------------
    # PowerPoint Copilot API
    # -------------------------------------------------------------------------
    def ppt_copy_clipboard(self, dsl_text: str) -> Dict[str, Any]:
        dsl_text = dsl_text.strip()
        if not dsl_text:
            return {"success": False, "error": "DSL is empty — nothing to copy."}
        try:
            refresh_dsl_theme_colors()
        except Exception:
            pass

        try:
            slides = parse_dsl_slides(dsl_text)
            shapes = slides[0] if slides else []
            if not shapes:
                return {"success": False, "error": "No shapes parsed from DSL."}

            self.ppt_connector.create_shapes_and_copy(shapes)
            return {"success": True, "message": "✓ Shapes copied to clipboard — switch to PowerPoint and press Ctrl+V."}
        except Exception as err:
            return {"success": False, "error": f"Clipboard copy failed: {err}"}

    def ppt_insert_current_slide(self, dsl_text: str) -> Dict[str, Any]:
        dsl_text = dsl_text.strip()
        if not dsl_text:
            return {"success": False, "error": "DSL is empty — nothing to insert."}
        try:
            refresh_dsl_theme_colors()
        except Exception:
            pass

        try:
            slides = parse_dsl_slides(dsl_text)
            shapes = slides[0] if slides else []
            if not shapes:
                return {"success": False, "error": "No shapes parsed from DSL."}

            self.ppt_connector.create_on_current_slide(shapes)
            return {"success": True, "message": "✓ Shapes successfully inserted on active PowerPoint slide!"}
        except Exception as err:
            return {"success": False, "error": f"Insert failed: {err}"}

    def ppt_create_full_slides(self, dsl_text: str) -> Dict[str, Any]:
        dsl_text = dsl_text.strip()
        if not dsl_text:
            return {"success": False, "error": "DSL is empty — nothing to build."}
        try:
            refresh_dsl_theme_colors()
        except Exception:
            pass

        try:
            slides = parse_dsl_slides(dsl_text)
            total_shapes = sum(len(s) for s in slides)
            if total_shapes == 0:
                return {"success": False, "error": "No valid shapes found to create."}

            self.ppt_connector.create_on_new_slide(slides)
            return {"success": True, "message": f"✓ Created {len(slides)} slide(s) ({total_shapes} shapes total) successfully!"}
        except Exception as err:
            return {"success": False, "error": f"Slide creation failed: {err}"}

    # -------------------------------------------------------------------------
    # Word Copilot API
    # -------------------------------------------------------------------------
    def word_build_and_open(self, dsl_text: str) -> Dict[str, Any]:
        dsl_text = dsl_text.strip()
        if not dsl_text:
            return {"success": False, "error": "DSL content is empty."}
        try:
            refresh_word_theme_colors()
        except Exception:
            pass

        try:
            pages = parse_dsl_pages(dsl_text)
            if not any(pages):
                return {"success": False, "error": "Nothing to build."}
            self.word_connector.build_and_open(pages)
            return {"success": True, "message": "✓ Document built and opened in Word!"}
        except Exception as err:
            return {"success": False, "error": f"Build document failed: {err}"}

    def word_insert_at_cursor(self, dsl_text: str) -> Dict[str, Any]:
        dsl_text = dsl_text.strip()
        if not dsl_text:
            return {"success": False, "error": "DSL content is empty."}
        try:
            refresh_word_theme_colors()
        except Exception:
            pass

        try:
            pages = parse_dsl_pages(dsl_text)
            total_elements = sum(len(p) for p in pages)
            if total_elements == 0:
                return {"success": False, "error": "No elements parsed from DSL."}

            self.word_connector.insert_at_cursor(pages)
            return {"success": True, "message": f"✓ Content inserted ({total_elements} elements) at cursor in Word!"}
        except Exception as err:
            return {"success": False, "error": f"Word cursor insertion failed: {err}"}

    def word_apply_edits(self, dsl_text: str) -> Dict[str, Any]:
        dsl_text = dsl_text.strip()
        if not dsl_text:
            return {"success": False, "error": "Edit instructions are empty."}
        try:
            refresh_word_theme_colors()
        except Exception:
            pass

        try:
            ops = parse_edit_dsl(dsl_text)
            if not ops:
                return {"success": False, "error": "No valid edit instructions parsed (use replace, insert_before, etc.)"}

            self.word_editor.apply(ops)
            return {"success": True, "message": f"✓ Applied {len(ops)} edit(s) to the document!"}
        except Exception as err:
            return {"success": False, "error": f"Apply edits failed: {err}"}

    def word_extract_dsl(self) -> Dict[str, Any]:
        try:
            dsl = self.word_extractor.extract()
            return {"success": True, "dsl": dsl, "message": "✓ Successfully extracted document structure as DSL!"}
        except Exception as err:
            return {"success": False, "error": f"Extraction failed: {err}"}

    # -------------------------------------------------------------------------
    # Excel Copilot API
    # -------------------------------------------------------------------------
    def excel_connect_active(self) -> Dict[str, Any]:
        """Connect directly to the currently active workbook in Microsoft Excel."""
        return self._analyze_excel_path(None)

    def excel_open_file(self) -> Dict[str, Any]:
        """Prompt file picker, open workbook in Microsoft Excel via COM, and analyze."""
        if not self._window:
            return {"success": False, "error": "Window not initialized."}
        file_types = ("Excel Workbooks (*.xlsx;*.xlsm)", "All files (*.*)")
        res = self._window.create_file_dialog(webview.OPEN_DIALOG, allow_multiple=False, file_types=file_types)
        if not res or len(res) == 0:
            return {"success": False, "cancelled": True}
        file_path = res[0]
        return self._analyze_excel_path(file_path)

    def _analyze_excel_path(self, file_path: Optional[str] = None) -> Dict[str, Any]:
        try:
            model = WorkbookAnalyzer.analyze(file_path)
            self.excel_current_path = model.file_path or model.filename
            self.excel_current_model = model

            # Extract metrics & sheets
            sheets_data = []
            formula_count = 0
            table_count = 0

            for sheet in model.worksheets:
                formula_count += sheet.formulas_count
                table_count += len(sheet.tables)
                sheets_data.append({
                    "name": sheet.name,
                    "row_count": sheet.max_row,
                    "col_count": sheet.max_column,
                    "has_headers": True if sheet.tables else False,
                    "headers": [c.name for t in sheet.tables for c in t.columns][:12],
                    "formulas_count": sheet.formulas_count,
                    "tables_count": len(sheet.tables),
                })

            # LLM Prompt Context
            context_text = WorkbookAnalyzer.generate_system_prompt(model)

            return {
                "success": True,
                "file_path": self.excel_current_path,
                "file_name": model.filename,
                "sheet_count": len(model.worksheets),
                "table_count": table_count,
                "formula_count": formula_count,
                "sheets": sheets_data,
                "context_text": context_text,
                "message": f"✓ Connected to active workbook: {model.filename}",
            }
        except Exception as err:
            traceback.print_exc()
            return {"success": False, "error": f"Failed to analyze active Excel workbook: {err}"}

    def excel_execute_protocol(self, protocol_json_str: str, create_backup: bool = True) -> Dict[str, Any]:
        try:
            protocol = ActionParser.parse_response(protocol_json_str)
            result, updated_model = self.excel_executor.execute(
                protocol=protocol,
                file_path_or_wb=self.excel_current_path,
                model=self.excel_current_model,
                create_backup=create_backup,
            )
            if updated_model:
                self.excel_current_model = updated_model

            log_messages = [f"[{d.get('status', 'info').upper()}] {d.get('message')}" for d in result.details]
            if result.errors:
                log_messages.extend([f"[ERROR] {e}" for e in result.errors])
            if result.warnings:
                log_messages.extend([f"[WARNING] {w}" for w in result.warnings])

            is_ok = result.status.value in ("success", "partial_success")
            return {
                "success": is_ok,
                "executed_actions": result.actions_executed,
                "total_actions": len(protocol.actions),
                "logs": log_messages,
                "message": f"{'✓' if is_ok else '⚠'} Execution finished: {result.actions_executed}/{len(protocol.actions)} action(s) succeeded.",
            }
        except Exception as err:
            return {"success": False, "error": f"Execution error: {err}"}

    # -------------------------------------------------------------------------
    # CV Copilot API
    # -------------------------------------------------------------------------
    def cv_run_audit(self, cv_json_str: str) -> Dict[str, Any]:
        try:
            data = json.loads(cv_json_str)
            audit_report = run_dq_audit(data)
            
            # Format report for UI
            issues = []
            for issue in audit_report.issues:
                issues.append({
                    "severity": issue.severity.value,
                    "field": issue.field,
                    "message": issue.message,
                    "code": issue.code,
                })

            summary = {
                "dq_score": audit_report.overall_score,
                "passed": audit_report.passed,
                "total_checks": audit_report.total_checks,
                "passed_checks": audit_report.passed_checks,
                "issues": issues,
                "candidate_name": f"{data.get('personal_information', {}).get('first_name', '')} {data.get('personal_information', {}).get('last_name', '')}".strip() or "Unnamed",
                "experience_count": len(data.get("work_experience", [])),
                "skills_count": len(data.get("skills", [])),
            }
            return {"success": True, "audit": summary}
        except Exception as err:
            return {"success": False, "error": f"Audit failed: {err}"}

    def cv_generate_docx(self, cv_json_str: str) -> Dict[str, Any]:
        try:
            data = json.loads(cv_json_str)
            import tempfile
            first_name = data.get("personal_information", {}).get("first_name", "Candidate")
            last_name = data.get("personal_information", {}).get("last_name", "CV")
            filename = f"CV_{first_name}_{last_name}.docx".replace(" ", "_")
            out_path = os.path.join(tempfile.gettempdir(), filename)

            generate_cv(data, out_path)
            try:
                os.startfile(out_path)
            except Exception:
                pass
            return {"success": True, "path": out_path, "message": f"✓ Europass CV generated and opened: {filename}"}
        except Exception as err:
            return {"success": False, "error": f"Word CV generation failed: {err}"}

    # -------------------------------------------------------------------------
    # Python Copilot API
    # -------------------------------------------------------------------------
    def python_run_code(self, code: str, persist: Optional[bool] = None) -> Dict[str, Any]:
        """Execute pasted Python code inside the sandbox workspace."""
        try:
            return self.python_runner.run_code(code=code, persist=persist)
        except Exception as err:
            return {
                "success": False,
                "error": f"Run failed: {err}",
                "stdout": "",
                "stderr": str(err),
                "exit_code": -1,
                "duration_ms": 0,
                "produced_files": [],
                "all_files": [],
            }

    def python_get_folder_context(self) -> Dict[str, Any]:
        """Retrieve structured markdown list of files in the current folder for LLM prompts."""
        try:
            return self.python_runner.get_folder_context_for_llm()
        except Exception as err:
            return {"success": False, "error": str(err)}

    def python_open_folder(self) -> Dict[str, Any]:
        """Open the active sandbox folder in Windows Explorer."""
        try:
            return self.python_runner.open_in_explorer()
        except Exception as err:
            return {"success": False, "error": str(err)}

    def python_open_file(self, filename: str) -> Dict[str, Any]:
        """Open a specific generated document or file in default application."""
        try:
            return self.python_runner.open_specific_file(filename)
        except Exception as err:
            return {"success": False, "error": str(err)}

    def python_clear_sandbox(self) -> Dict[str, Any]:
        """Clear the current sandbox files."""
        try:
            return self.python_runner.clear_sandbox()
        except Exception as err:
            return {"success": False, "error": str(err)}

    def python_set_persistence(self, persist: bool) -> Dict[str, Any]:
        """Toggle persistence mode on/off."""
        try:
            return self.python_runner.set_persistence(persist)
        except Exception as err:
            return {"success": False, "error": str(err)}

    def python_get_status(self) -> Dict[str, Any]:
        """Return current status: working folder, persistence flag, and file list."""
        try:
            return {
                "success": True,
                "folder_path": str(self.python_runner.current_dir),
                "persist": self.python_runner.persist_mode,
                "files": self.python_runner.list_files(),
            }
        except Exception as err:
            return {"success": False, "error": str(err)}

