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
    PPT_SAMPLES,
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
from copilots_app.services.word.sample import SAMPLE_DSL

# Excel services
from copilots_app.services.excel.analyzer.workbook_analyzer import WorkbookAnalyzer
from copilots_app.services.excel.executor.executor_main import ActionExecutor
from copilots_app.services.excel.protocol.action_parser import ActionParser
from copilots_app.services.excel.models.protocol import ActionProtocol

# CV services
from copilots_app.services.cv import load_sample_cv, run_dq_audit, generate_cv


WORD_SAMPLES = {
    "Complete Demo": SAMPLE_DSL,
    "Executive Report": """// Document setup
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
row="Automation Suite","Desktop Copilots unification with pywebview","✓ Production"
row="Data Quality","Automated compliance validation","✓ Verified"

br

h2 color=a4 | "Key Recommendations"
ul indent=1
  item | "Maintain strict deterministic data validation across all output formats."
  item | "Accelerate user enablement sessions with interactive cheatsheets."
  item | "Continuous integration testing for COM office automation hooks."
""",
    "Edit Mode Example": """// Target active document in Microsoft Word
edit target=active

// Search and replace or update headings
replace find="Draft" replace="Final Approved Version"
insert_before find="Key Recommendations" | "Note: All milestones reviewed by Executive Sponsor."
""",
}

SAMPLE_EXCEL_ACTION_JSON = """{
  "intent": "MODIFY_WORKBOOK",
  "version": "1.0",
  "actions": [
    {
      "action": "SET_CELL_VALUE",
      "sheet": "Summary",
      "cell": "B2",
      "value": "Updated by Copilot Suite"
    },
    {
      "action": "FORMAT_CELL",
      "sheet": "Summary",
      "range": "B2:D2",
      "bold": true,
      "fill_color": "E2EFDA",
      "font_color": "375623"
    }
  ]
}"""


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
    def get_ppt_samples(self) -> Dict[str, str]:
        return PPT_SAMPLES

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
    def get_word_samples(self) -> Dict[str, str]:
        return WORD_SAMPLES

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
    def get_excel_sample_json(self) -> str:
        return SAMPLE_EXCEL_ACTION_JSON

    def excel_open_file(self) -> Dict[str, Any]:
        if not self._window:
            return {"success": False, "error": "Window not initialized."}
        file_types = ("Excel Workbooks (*.xlsx;*.xlsm)", "All files (*.*)")
        res = self._window.create_file_dialog(webview.OPEN_DIALOG, allow_multiple=False, file_types=file_types)
        if not res or len(res) == 0:
            return {"success": False, "cancelled": True}
        file_path = res[0]
        return self._analyze_excel_path(file_path)

    def excel_open_demo(self) -> Dict[str, Any]:
        base_dir = os.path.dirname(os.path.abspath(__file__))
        demo_path = os.path.join(base_dir, "services", "excel", "sample_portfolio.xlsx")
        if not os.path.exists(demo_path):
            from copilots_app.services.excel.sample_data import create_sample_workbook
            try:
                create_sample_workbook(demo_path)
            except Exception as e:
                return {"success": False, "error": f"Could not create demo workbook: {e}"}
        return self._analyze_excel_path(demo_path)

    def _analyze_excel_path(self, file_path: str) -> Dict[str, Any]:
        try:
            analyzer = WorkbookAnalyzer(file_path)
            model = analyzer.analyze()
            self.excel_current_path = file_path
            self.excel_current_model = model

            # Extract metrics & sheets
            sheets_data = []
            formula_count = 0
            table_count = 0

            for sheet in model.sheets:
                formula_count += len(sheet.formulas)
                table_count += len(sheet.tables)
                sheets_data.append({
                    "name": sheet.name,
                    "row_count": sheet.row_count,
                    "col_count": sheet.col_count,
                    "has_headers": sheet.has_headers,
                    "headers": sheet.headers[:12],
                    "formulas_count": len(sheet.formulas),
                    "tables_count": len(sheet.tables),
                })

            # LLM Prompt Context
            from copilots_app.services.excel.analyzer.llm_context_generator import LLMContextGenerator
            ctx_gen = LLMContextGenerator(model)
            context_text = ctx_gen.generate_full_context()

            return {
                "success": True,
                "file_path": file_path,
                "file_name": os.path.basename(file_path),
                "sheet_count": len(model.sheets),
                "table_count": table_count,
                "formula_count": formula_count,
                "sheets": sheets_data,
                "context_text": context_text,
                "message": f"✓ Loaded and analyzed {os.path.basename(file_path)}",
            }
        except Exception as err:
            traceback.print_exc()
            return {"success": False, "error": f"Failed to analyze workbook: {err}"}

    def excel_execute_protocol(self, protocol_json_str: str, create_backup: bool = True) -> Dict[str, Any]:
        if not self.excel_current_path:
            return {"success": False, "error": "No workbook currently open. Please open a workbook first."}
        try:
            data = json.loads(protocol_json_str)
            protocol = ActionParser.parse_dict(data)
            status = self.excel_executor.execute(self.excel_current_path, protocol, create_backup=create_backup)
            
            log_messages = [log.message for log in status.logs]
            return {
                "success": status.success,
                "executed_actions": status.actions_completed,
                "total_actions": status.total_actions,
                "logs": log_messages,
                "message": f"{'✓' if status.success else '⚠'} Execution finished: {status.actions_completed}/{status.total_actions} actions succeeded.",
            }
        except Exception as err:
            return {"success": False, "error": f"Execution error: {err}"}

    # -------------------------------------------------------------------------
    # CV Copilot API
    # -------------------------------------------------------------------------
    def cv_get_sample_data(self) -> Dict[str, Any]:
        try:
            data = load_sample_cv()
            return {"success": True, "cv_data": data}
        except Exception as err:
            return {"success": False, "error": str(err)}

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
