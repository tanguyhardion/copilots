"""
Bridge API exposed to pywebview JavaScript frontend (window.pywebview.api).
Provides full access to PowerPoint and CV Copilot engines and System Prompts.
"""

import os
import json
import tempfile
from typing import Dict, Any, Optional
import webview

from copilots_app import __version__

from copilots_app.core.prompt_manager import PromptManager

# PowerPoint services
from copilots_app.services.powerpoint import (
    PowerPointConnector,
    parse_dsl_slides,
    refresh_dsl_theme_colors,
)

# CV services
from copilots_app.services.cv import run_dq_audit, generate_cv, generate_pptx_cv


class CopilotBridge:
    """Python API bridge passed to pywebview."""

    def __init__(self):
        self._window = None
        self.prompt_manager = PromptManager()
        self.ppt_connector = PowerPointConnector()

    def set_window(self, window):
        self._window = window

    def get_version(self) -> str:
        """Return the application version string."""
        return __version__

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
            render_shapes = [s for s in shapes if s.get("type") != "slide"]
            if not shapes:
                return {"success": False, "error": "No shapes parsed from DSL."}

            self.ppt_connector.create_shapes_and_copy(shapes)
            if render_shapes:
                return {"success": True, "message": "✓ Shapes copied to clipboard — switch to PowerPoint and press Ctrl+V."}
            return {"success": True, "message": "✓ Slide background updated on temporary slide (no drawable shapes to copy)."}
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
            render_shapes = [s for s in shapes if s.get("type") != "slide"]
            if not shapes:
                return {"success": False, "error": "No shapes parsed from DSL."}

            self.ppt_connector.create_on_current_slide(shapes)
            if render_shapes:
                return {"success": True, "message": "✓ Shapes successfully inserted on active PowerPoint slide!"}
            return {"success": True, "message": "✓ Active slide background color updated successfully!"}
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
            total_shapes = sum(len([shape for shape in s if shape.get("type") != "slide"]) for s in slides)
            total_directives = sum(len([shape for shape in s if shape.get("type") == "slide"]) for s in slides)
            if total_shapes == 0 and total_directives == 0:
                return {"success": False, "error": "No valid shapes found to create."}

            self.ppt_connector.create_on_new_slide(slides)
            if total_shapes > 0:
                return {"success": True, "message": f"✓ Created {len(slides)} slide(s) ({total_shapes} shapes total) successfully!"}
            return {"success": True, "message": f"✓ Created {len(slides)} slide(s) with background directives successfully!"}
        except Exception as err:
            return {"success": False, "error": f"Slide creation failed: {err}"}

    # -------------------------------------------------------------------------
    # CV Copilot API
    # -------------------------------------------------------------------------
    def cv_run_audit(self, cv_json_str: str) -> Dict[str, Any]:
        try:
            data = json.loads(cv_json_str)
            flags = run_dq_audit(data)
            
            # Format report for UI
            issues = []
            error_count = 0
            warning_count = 0
            for flag in flags:
                sev = (flag.get("severity") or "INFO").upper()
                if sev == "ERROR":
                    error_count += 1
                elif sev == "WARNING":
                    warning_count += 1

                issues.append({
                    "severity": sev,
                    "field": flag.get("field") or flag.get("section") or "General",
                    "message": flag.get("message", ""),
                    "code": flag.get("rule_code", ""),
                })

            total_rules = 12
            passed = error_count == 0
            dq_score = max(0, min(100, 100 - (error_count * 20 + warning_count * 5)))

            pi = data.get("personal_info") or data.get("personal_information") or {}
            first_name = pi.get("first_name", "")
            last_name = pi.get("last_name", "")
            candidate_name = f"{first_name} {last_name}".strip() or "Unnamed"

            summary = {
                "dq_score": dq_score,
                "passed": passed,
                "total_checks": total_rules,
                "passed_checks": max(0, total_rules - error_count),
                "issues": issues,
                "candidate_name": candidate_name,
                "experience_count": len(data.get("work_experience", []) or data.get("project_experience", [])),
                "skills_count": len(data.get("skills", []) or data.get("personal_skills", {}).get("communication", [])),
            }
            return {"success": True, "audit": summary}
        except Exception as err:
            return {"success": False, "error": f"Audit failed: {err}"}

    def _resolve_cv_output_path(self, filename: str, output_dir: Optional[str] = None) -> str:
        target_dir = (output_dir or "").strip() or tempfile.gettempdir()
        target_dir = os.path.abspath(os.path.expanduser(target_dir))
        os.makedirs(target_dir, exist_ok=True)
        return os.path.join(target_dir, filename)

    def cv_select_output_folder(self) -> Dict[str, Any]:
        if not self._window:
            return {"success": False, "error": "Window not initialized."}
        res = self._window.create_file_dialog(webview.FOLDER_DIALOG, allow_multiple=False)
        if not res or len(res) == 0:
            return {"success": False, "cancelled": True}
        return {"success": True, "path": res[0]}

    def cv_generate_docx(
        self,
        cv_json_str: str,
        language: str = "en",
        sections: Optional[Dict[str, bool]] = None,
        output_dir: Optional[str] = None,
    ) -> Dict[str, Any]:
        try:
            data = json.loads(cv_json_str)
            pi = data.get("personal_info") or data.get("personal_information") or {}
            first_name = pi.get("first_name", "Candidate")
            last_name = pi.get("last_name", "CV")
            filename = f"CV_{first_name}_{last_name}.docx".replace(" ", "_")
            out_path = self._resolve_cv_output_path(filename, output_dir)

            generate_cv(data, out_path, language=language, sections=sections)
            try:
                os.startfile(out_path)
            except Exception:
                pass
            return {"success": True, "path": out_path, "message": f"✓ Europass CV generated and opened: {filename}"}
        except Exception as err:
            return {"success": False, "error": f"Word CV generation failed: {err}"}

    def cv_generate_pptx(
        self,
        cv_json_str: str,
        language: str = "en",
        sections: Optional[Dict[str, bool]] = None,
        output_dir: Optional[str] = None,
    ) -> Dict[str, Any]:
        try:
            data = json.loads(cv_json_str)
            pi = data.get("personal_info") or data.get("personal_information") or {}
            first_name = pi.get("first_name", "Candidate")
            last_name = pi.get("last_name", "CV")
            filename = f"CV_{first_name}_{last_name}_1Slide.pptx".replace(" ", "_")
            out_path = self._resolve_cv_output_path(filename, output_dir)

            generate_pptx_cv(data, out_path, language=language, sections=sections)
            try:
                os.startfile(out_path)
            except Exception:
                pass
            return {"success": True, "path": out_path, "message": f"✓ 1-Slide Executive PowerPoint CV generated and opened: {filename}"}
        except Exception as err:
            return {"success": False, "error": f"PowerPoint CV generation failed: {err}"}
