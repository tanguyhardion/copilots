"""
pptx_generator.py
Generates a 1-slide Executive Proposal CV in PowerPoint (.pptx) format
by precisely modifying template.pptx in-place to preserve 100% of the
original layout, geometry, positioning, typography, colors, and styling.
"""

from __future__ import annotations

import os
import sys
from pathlib import Path
from typing import Any, Dict, List, Optional
import pptx
from pptx.dml.color import RGBColor
from pptx.util import Pt
from pptx.enum.text import PP_ALIGN

from copilots_app.services.cv.i18n import (
    PPTX_SECTION_DEFAULTS,
    get_cv_strings,
    normalize_sections,
)

FONT_NAME = "Calibri"
COLOR_DARK = RGBColor(36, 36, 36)  # RGB(36, 36, 36) matching template.pptx runs
COLOR_BLACK = RGBColor(0, 0, 0)

def _resolve_template_path(template_path: Optional[str] = None) -> str:
    """Find the template.pptx file across dev workspaces and frozen builds."""
    if template_path and os.path.exists(template_path):
        return template_path

    # Check PyInstaller bundle
    if getattr(sys, "frozen", False) and hasattr(sys, "_MEIPASS"):
        meipass_tpl = Path(sys._MEIPASS) / "assets" / "templates" / "cv_template.pptx"
        if meipass_tpl.exists():
            return str(meipass_tpl)

    # Check project assets folder
    base_dir = Path(__file__).resolve().parent.parent.parent.parent
    candidates = [
        base_dir / "assets" / "templates" / "cv_template.pptx",
        base_dir / ".old" / "cv-copilot" / "template.pptx",
        Path(__file__).resolve().parent / "template.pptx",
    ]
    for c in candidates:
        if c.exists():
            return str(c)

    raise FileNotFoundError("Executive CV template.pptx could not be found.")


def _fmt_date_part(ym: str, strings: Dict[str, Any]) -> str:
    """Format YYYY-MM to MMM. YYY (e.g. 2021-01 -> Jan. 2021) or keep 'Present'."""
    if not ym:
        return ""
    if ym.lower() == "present":
        return strings["present"]
    parts = ym.split("-")
    if len(parts) == 2:
        try:
            m_idx = int(parts[1]) - 1
            if 0 <= m_idx < 12:
                return f"{strings['pptx_months'][m_idx]} {parts[0]}"
        except ValueError:
            pass
    return ym


def _fmt_date_range(date_from: str, date_to: str, strings: Dict[str, Any]) -> str:
    """Format date range: 'Jan. 2021 – Present'."""
    f = _fmt_date_part(date_from, strings)
    t = _fmt_date_part(date_to, strings)
    if f and t:
        return f"{f} – {t}"
    return f or t or ""


def _clear_text_frame(text_frame) -> None:
    for paragraph in text_frame.paragraphs:
        paragraph.text = ""


def generate_pptx_cv(
    cv_json: Dict[str, Any],
    output_path: str,
    template_path: Optional[str] = None,
    language: str = "en",
    sections: Optional[Dict[str, bool]] = None,
) -> str:
    """
    Populate template.pptx with candidate information while preserving exact formatting.
    """
    strings = get_cv_strings(language)
    selected_sections = normalize_sections(sections, PPTX_SECTION_DEFAULTS)
    resolved_template = _resolve_template_path(template_path)
    prs = pptx.Presentation(resolved_template)
    slide = prs.slides[0]

    # Map personal info
    pi = cv_json.get("personal_info") or cv_json.get("personal_information") or {}
    first_name = (pi.get("first_name") or "").strip()
    last_name = (pi.get("last_name") or "").strip()
    full_name = f"{first_name} {last_name}".strip() or strings["candidate_name"]

    # Tender / proposed role
    tender_info = cv_json.get("tender_info") or {}
    proposed_role = (tender_info.get("proposed_role") or pi.get("grade") or pi.get("role") or "").strip()

    # Team / Industry
    team_or_industry = (pi.get("team") or pi.get("industry") or tender_info.get("industry") or "").strip()

    # Contact
    phone = (pi.get("phone") or "").strip()
    email = (pi.get("email") or "").strip()

    # -------------------------------------------------------------------------
    # 1. Candidate Header & Contact Info (Text Placeholder 2)
    # -------------------------------------------------------------------------
    tp_shapes = [s for s in slide.shapes if s.name == "Text Placeholder 2"]
    if tp_shapes:
        tf = tp_shapes[0].text_frame
        if not selected_sections["header_contact"]:
            _clear_text_frame(tf)
        else:
            # Line 0: Full name
            if len(tf.paragraphs) > 0:
                p0 = tf.paragraphs[0]
                p0.text = full_name
                if p0.runs:
                    p0.runs[0].font.name = FONT_NAME
                    p0.runs[0].font.size = Pt(11)

            # Line 1: Grade / Rank / Role
            if len(tf.paragraphs) > 1:
                p1 = tf.paragraphs[1]
                p1.text = proposed_role if proposed_role else ""
                if p1.runs:
                    p1.runs[0].font.name = FONT_NAME
                    p1.runs[0].font.size = Pt(11)

            # Line 2: Team / Industry
            if len(tf.paragraphs) > 2:
                p2 = tf.paragraphs[2]
                p2.text = team_or_industry if team_or_industry else ""
                if p2.runs:
                    p2.runs[0].font.name = FONT_NAME
                    p2.runs[0].font.size = Pt(11)

            # Line 3: Mobile
            if len(tf.paragraphs) > 3:
                p3 = tf.paragraphs[3]
                p3.text = f"{strings['mobile']}: {phone}" if phone else ""
                if p3.runs:
                    p3.runs[0].font.name = FONT_NAME
                    p3.runs[0].font.size = Pt(11)

            # Line 4: Email
            if len(tf.paragraphs) > 4:
                p4 = tf.paragraphs[4]
                p4.text = f"{strings['email']}: {email}" if email else ""
                if p4.runs:
                    p4.runs[0].font.name = FONT_NAME
                    p4.runs[0].font.size = Pt(11)

    # -------------------------------------------------------------------------
    # 2. Profile Summary (TextBox 7)
    # -------------------------------------------------------------------------
    tb_shapes = [s for s in slide.shapes if s.name == "TextBox 7"]
    if tb_shapes:
        tf_profile = tb_shapes[0].text_frame
        if not selected_sections["profile"]:
            _clear_text_frame(tf_profile)
        else:
            profile_text = (cv_json.get("profile") or "").strip()
            if not profile_text:
                profile_text = (cv_json.get("summary") or "").strip()

            if profile_text and tf_profile.paragraphs:
                p_prof = tf_profile.paragraphs[0]
                p_prof.text = profile_text
                if p_prof.runs:
                    p_prof.runs[0].font.name = FONT_NAME
                    p_prof.runs[0].font.size = Pt(11)

    # -------------------------------------------------------------------------
    # 3. Subject Matter Expertise (Table 72 - 2 rows x 3 columns = 6 skills)
    # -------------------------------------------------------------------------
    t72_shapes = [s for s in slide.shapes if s.name == "Table 72" and s.has_table]
    if t72_shapes:
        table_skills = t72_shapes[0].table
        skills_list: List[str] = []

        if selected_sections["skills"]:
            # Try personal_skills categorized list first
            ps = cv_json.get("personal_skills") or {}
            if isinstance(ps, dict):
                for cat in ["communication", "organisational_managerial", "computer_skills"]:
                    for s in ps.get(cat, []):
                        if isinstance(s, str) and s.strip() and s.strip() not in skills_list:
                            skills_list.append(s.strip())

            # Fallback to plain skills array if needed
            if not skills_list:
                raw_skills = cv_json.get("skills") or []
                for s in raw_skills:
                    if isinstance(s, str) and s.strip():
                        skills_list.append(s.strip())
                    elif isinstance(s, dict) and s.get("name"):
                        skills_list.append(s["name"].strip())

        skill_idx = 0
        for r in range(len(table_skills.rows)):
            for c in range(len(table_skills.columns)):
                cell = table_skills.cell(r, c)
                val = skills_list[skill_idx] if skill_idx < len(skills_list) else ""
                cell.text = val
                # Re-apply bold 12pt Calibri to runs
                for p in cell.text_frame.paragraphs:
                    for run in p.runs:
                        run.font.name = FONT_NAME
                        run.font.size = Pt(12)
                        run.font.bold = True
                skill_idx += 1

    # -------------------------------------------------------------------------
    # 4. Relevant Experience (Table 71 - 5 rows x 2 columns)
    # -------------------------------------------------------------------------
    t71_shapes = [s for s in slide.shapes if s.name == "Table 71" and s.has_table]
    if t71_shapes:
        table_exp = t71_shapes[0].table
        projects = []
        if selected_sections["relevant_experience"]:
            projects = cv_json.get("project_experience") or []
            if not projects:
                # Fall back to work_experience if project_experience not provided
                projects = cv_json.get("work_experience") or []

        max_rows = len(table_exp.rows)

        for row_idx in range(max_rows):
            cell_date = table_exp.cell(row_idx, 0)
            cell_desc = table_exp.cell(row_idx, 1)

            if row_idx < len(projects):
                proj = projects[row_idx]
                d_from = proj.get("date_from", "")
                d_to = proj.get("date_to", "")
                date_label = _fmt_date_range(d_from, d_to, strings)

                # Set date in col 0
                cell_date.text = date_label
                for p in cell_date.text_frame.paragraphs:
                    for run in p.runs:
                        run.font.name = FONT_NAME
                        run.font.size = Pt(11)
                        run.font.bold = False

                # Build description header in col 1
                # Format: "<Sector/Client> – <Project Title>"
                client = proj.get("client") or proj.get("sector") or proj.get("organisation") or ""
                title = proj.get("project_title") or proj.get("job_title") or proj.get("role") or ""
                if client and title:
                    header_line = f"{client} – {title}"
                else:
                    header_line = title or client or strings["project"]

                # Clear and populate paragraphs
                cell_desc.text = ""  # Clears text frame and leaves 1 empty paragraph
                p0 = cell_desc.text_frame.paragraphs[0]
                p0.text = header_line
                p0.space_after = Pt(0)
                p0.space_before = Pt(0)
                for run in p0.runs:
                    run.font.name = FONT_NAME
                    run.font.size = Pt(11)
                    run.font.bold = True
                    run.font.color.rgb = COLOR_DARK

                # Bullets
                bullets = proj.get("bullets") or []
                if isinstance(bullets, str):
                    bullets = [bullets]
                elif not bullets and proj.get("responsibilities"):
                    bullets = [proj["responsibilities"]]

                # Keep up to 3 most impactful bullets to avoid table overflow
                for b_text in bullets[:3]:
                    if not b_text or not b_text.strip():
                        continue
                    p_b = cell_desc.text_frame.add_paragraph()
                    p_b.text = b_text.strip()
                    p_b.space_after = Pt(0)
                    p_b.space_before = Pt(0)
                    for run in p_b.runs:
                        run.font.name = FONT_NAME
                        run.font.size = Pt(11)
                        run.font.bold = False
                        run.font.color.rgb = COLOR_DARK
            else:
                # Clear unused rows cleanly so template placeholder text is not visible
                cell_date.text = ""
                cell_desc.text = ""

    # Ensure target directory exists
    os.makedirs(os.path.dirname(os.path.abspath(output_path)), exist_ok=True)
    prs.save(output_path)
    return output_path
