"""
pptx_generator.py
Generates a 1-slide Executive Proposal CV in PowerPoint (.pptx) format
using PowerPoint COM Automation (win32com) to modify template.pptx
and preserve 100% of the original layout, typography, and styling.
"""

from __future__ import annotations

import os
import sys
import shutil
from pathlib import Path
from typing import Any, Dict, List, Optional

import pythoncom
import win32com.client

from copilots_app.services.cv.i18n import (
    PPTX_SECTION_DEFAULTS,
    get_cv_strings,
    normalize_sections,
)

FONT_NAME = "Calibri"
COLOR_DARK_RGB = 0x242424  # RGB(36, 36, 36) in BGR/RGB integer


def _rgb_to_int(r: int, g: int, b: int) -> int:
    """Convert RGB (0-255) to Windows BGR COLORREF integer used by PowerPoint COM."""
    return r | (g << 8) | (b << 16)


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


def _resolve_photo_placeholder_path() -> Optional[str]:
    """Find the photo_placeholder.jpg asset across dev workspaces and frozen builds."""
    # Check PyInstaller bundle
    if getattr(sys, "frozen", False) and hasattr(sys, "_MEIPASS"):
        p = Path(sys._MEIPASS) / "assets" / "templates" / "photo_placeholder.jpg"
        if p.exists():
            return str(p)

    base_dir = Path(__file__).resolve().parent.parent.parent.parent
    candidates = [
        base_dir / "assets" / "templates" / "photo_placeholder.jpg",
        Path(__file__).resolve().parent / "photo_placeholder.jpg",
    ]
    for c in candidates:
        if c.exists():
            return str(c)
    return None


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


def _clear_text_frame(shape) -> None:
    """Clear text in a COM shape text frame."""
    try:
        shape.TextFrame.TextRange.Text = ""
    except Exception:
        pass


def generate_pptx_cv(
    cv_json: Dict[str, Any],
    output_path: str,
    template_path: Optional[str] = None,
    language: str = "en",
    sections: Optional[Dict[str, bool]] = None,
) -> str:
    """
    Populate template.pptx with candidate information using PowerPoint COM automation.
    """
    strings = get_cv_strings(language)
    selected_sections = normalize_sections(sections, PPTX_SECTION_DEFAULTS)
    resolved_template = _resolve_template_path(template_path)

    abs_output_path = os.path.abspath(output_path)
    os.makedirs(os.path.dirname(abs_output_path), exist_ok=True)

    # Make a copy of template to target output path
    shutil.copy2(resolved_template, abs_output_path)

    pythoncom.CoInitialize()
    ppt_app = None
    pres = None
    created_app = False

    try:
        try:
            ppt_app = win32com.client.GetActiveObject("PowerPoint.Application")
        except Exception:
            ppt_app = win32com.client.Dispatch("PowerPoint.Application")
            created_app = True

        # Open presentation
        pres = ppt_app.Presentations.Open(
            FileName=abs_output_path,
            ReadOnly=False,
            Untitled=False,
            WithWindow=False,
        )
        slide = pres.Slides(1)

        # Index shapes by name
        shapes_by_name: Dict[str, Any] = {}
        for i in range(1, slide.Shapes.Count + 1):
            s = slide.Shapes(i)
            shapes_by_name[s.Name] = s

        # -------------------------------------------------------------------------
        # Style fixes applied to shapes
        # -------------------------------------------------------------------------
        # Remove outline from "Text Placeholder 2"
        if "Text Placeholder 2" in shapes_by_name:
            try:
                shapes_by_name["Text Placeholder 2"].Line.Visible = False
            except Exception:
                pass

        # Remove outline + force black text on "Rectangle 3" (Profile Summary label)
        if "Rectangle 3" in shapes_by_name:
            try:
                s = shapes_by_name["Rectangle 3"]
                s.Line.Visible = False
                s.TextFrame.TextRange.Font.Color.RGB = _rgb_to_int(0, 0, 0)
            except Exception:
                pass

        # Remove outline from "TextBox 7" (Profile Summary content)
        if "TextBox 7" in shapes_by_name:
            try:
                shapes_by_name["TextBox 7"].Line.Visible = False
            except Exception:
                pass

        # Remove outline + force black text on "Rectangle 57" (Relevant Experience label)
        if "Rectangle 57" in shapes_by_name:
            try:
                s = shapes_by_name["Rectangle 57"]
                s.Line.Visible = False
                s.TextFrame.TextRange.Font.Color.RGB = _rgb_to_int(0, 0, 0)
            except Exception:
                pass

        # Remove outline on "Rectangle 49" (Subject Matter Expertise label)
        if "Rectangle 49" in shapes_by_name:
            try:
                shapes_by_name["Rectangle 49"].Line.Visible = False
            except Exception:
                pass

        # Remove cell shading and borders from Table 71 (Relevant Experience)
        if "Table 71" in shapes_by_name and shapes_by_name["Table 71"].HasTable:
            t71 = shapes_by_name["Table 71"].Table
            for r in range(1, t71.Rows.Count + 1):
                for c in range(1, t71.Columns.Count + 1):
                    try:
                        cell_shape = t71.Cell(r, c).Shape
                        cell_shape.Fill.Visible = False
                        cell_shape.Line.Visible = False
                    except Exception:
                        pass

        # Replace "Picture 4" placeholder rectangle with real image if present
        photo_placeholder_path = _resolve_photo_placeholder_path()
        if "Picture 4" in shapes_by_name and photo_placeholder_path:
            old_shape = shapes_by_name["Picture 4"]
            try:
                left = old_shape.Left
                top = old_shape.Top
                width = old_shape.Width
                height = old_shape.Height
                slide.Shapes.AddPicture(
                    FileName=os.path.abspath(photo_placeholder_path),
                    LinkToFile=False,
                    SaveWithDocument=True,
                    Left=left,
                    Top=top,
                    Width=width,
                    Height=height,
                )
                old_shape.Delete()
            except Exception:
                pass

        # -------------------------------------------------------------------------
        # Map personal info
        pi = cv_json.get("personal_info") or cv_json.get("personal_information") or {}
        first_name = (pi.get("first_name") or "").strip()
        last_name = (pi.get("last_name") or "").strip()
        full_name = f"{first_name} {last_name}".strip() or strings["candidate_name"]

        tender_info = cv_json.get("tender_info") or {}
        proposed_role = (tender_info.get("proposed_role") or pi.get("grade") or pi.get("role") or "").strip()
        team_or_industry = (pi.get("team") or pi.get("industry") or tender_info.get("industry") or "").strip()
        phone = (pi.get("phone") or "").strip()
        email = (pi.get("email") or "").strip()

        # -------------------------------------------------------------------------
        # 1. Candidate Header & Contact Info (Text Placeholder 2)
        # -------------------------------------------------------------------------
        if "Text Placeholder 2" in shapes_by_name:
            tp_shape = shapes_by_name["Text Placeholder 2"]
            if not selected_sections["header_contact"]:
                _clear_text_frame(tp_shape)
            else:
                tf = tp_shape.TextFrame
                lines = [
                    full_name,
                    proposed_role if proposed_role else "",
                    team_or_industry if team_or_industry else "",
                    f"{strings['mobile']}: {phone}" if phone else "",
                    f"{strings['email']}: {email}" if email else "",
                ]
                # Set lines
                tf.TextRange.Text = "\r".join(lines)
                for p_idx in range(1, tf.TextRange.Paragraphs().Count + 1):
                    p = tf.TextRange.Paragraphs(p_idx)
                    p.Font.Name = FONT_NAME
                    p.Font.Size = 11

        # -------------------------------------------------------------------------
        # 2. Profile Summary (TextBox 7)
        # -------------------------------------------------------------------------
        if "TextBox 7" in shapes_by_name:
            tb_shape = shapes_by_name["TextBox 7"]
            if not selected_sections["profile"]:
                _clear_text_frame(tb_shape)
            else:
                profile_text = (cv_json.get("profile") or "").strip()
                if not profile_text:
                    profile_text = (cv_json.get("summary") or "").strip()
                tf_profile = tb_shape.TextFrame
                tf_profile.TextRange.Text = profile_text
                tf_profile.TextRange.Font.Name = FONT_NAME
                tf_profile.TextRange.Font.Size = 11

        # -------------------------------------------------------------------------
        # 3. Subject Matter Expertise (Table 72 - 2 rows x 3 columns = 6 skills)
        # -------------------------------------------------------------------------
        if "Table 72" in shapes_by_name and shapes_by_name["Table 72"].HasTable:
            table_skills = shapes_by_name["Table 72"].Table
            skills_list: List[str] = []

            if selected_sections["skills"]:
                ps = cv_json.get("personal_skills") or {}
                if isinstance(ps, dict):
                    for cat in ["communication", "organisational_managerial", "computer_skills"]:
                        for s in ps.get(cat, []):
                            if isinstance(s, str) and s.strip() and s.strip() not in skills_list:
                                skills_list.append(s.strip())

                if not skills_list:
                    raw_skills = cv_json.get("skills") or []
                    for s in raw_skills:
                        if isinstance(s, str) and s.strip():
                            skills_list.append(s.strip())
                        elif isinstance(s, dict) and s.get("name"):
                            skills_list.append(s["name"].strip())

            skill_idx = 0
            gray_bgr = _rgb_to_int(232, 232, 232)  # #E8E8E8

            for r in range(1, table_skills.Rows.Count + 1):
                for c in range(1, table_skills.Columns.Count + 1):
                    cell = table_skills.Cell(r, c)
                    cell_shape = cell.Shape
                    val = skills_list[skill_idx] if skill_idx < len(skills_list) else ""
                    cell_shape.TextFrame.TextRange.Text = val
                    cell_shape.TextFrame.TextRange.Font.Name = FONT_NAME
                    cell_shape.TextFrame.TextRange.Font.Size = 12
                    cell_shape.TextFrame.TextRange.Font.Bold = True

                    # Center alignment
                    for pi in range(1, cell_shape.TextFrame.TextRange.Paragraphs().Count + 1):
                        cell_shape.TextFrame.TextRange.Paragraphs(pi).ParagraphFormat.Alignment = 2  # ppAlignCenter

                    # Shading
                    cell_shape.Fill.Solid()
                    cell_shape.Fill.ForeColor.RGB = gray_bgr
                    cell_shape.Line.Visible = False

                    # Margins
                    cell_shape.TextFrame.MarginLeft = 1
                    cell_shape.TextFrame.MarginRight = 1
                    cell_shape.TextFrame.MarginTop = 1
                    cell_shape.TextFrame.MarginBottom = 1

                    skill_idx += 1

        # -------------------------------------------------------------------------
        # 4. Relevant Experience (Table 71 - 5 rows x 2 columns)
        # -------------------------------------------------------------------------
        if "Table 71" in shapes_by_name and shapes_by_name["Table 71"].HasTable:
            table_exp = shapes_by_name["Table 71"].Table
            projects = []
            if selected_sections["relevant_experience"]:
                projects = cv_json.get("project_experience") or []
                if not projects:
                    projects = cv_json.get("work_experience") or []

            max_rows = table_exp.Rows.Count

            dark_bgr = _rgb_to_int(36, 36, 36)

            for row_idx in range(max_rows):
                cell_date = table_exp.Cell(row_idx + 1, 1).Shape
                cell_desc = table_exp.Cell(row_idx + 1, 2).Shape

                # Clear borders & fill
                cell_date.Fill.Visible = False
                cell_date.Line.Visible = False
                cell_desc.Fill.Visible = False
                cell_desc.Line.Visible = False

                if row_idx < len(projects):
                    proj = projects[row_idx]
                    d_from = proj.get("date_from", "")
                    d_to = proj.get("date_to", "")
                    date_label = _fmt_date_range(d_from, d_to, strings)

                    # Set date in col 1
                    cell_date.TextFrame.TextRange.Text = date_label
                    cell_date.TextFrame.TextRange.Font.Name = FONT_NAME
                    cell_date.TextFrame.TextRange.Font.Size = 11
                    cell_date.TextFrame.TextRange.Font.Bold = False

                    # Build description header in col 2
                    client = proj.get("client") or proj.get("sector") or proj.get("organisation") or ""
                    title = proj.get("project_title") or proj.get("job_title") or proj.get("role") or ""
                    if client and title:
                        header_line = f"{client} – {title}"
                    else:
                        header_line = title or client or strings["project"]

                    bullets = proj.get("bullets") or []
                    if isinstance(bullets, str):
                        bullets = [bullets]
                    elif not bullets and proj.get("responsibilities"):
                        bullets = [proj["responsibilities"]]

                    valid_bullets = [b.strip() for b in bullets[:3] if b and b.strip()]

                    # Format header + bullets
                    tf_desc = cell_desc.TextFrame
                    all_lines = [header_line] + [f"• {b}" for b in valid_bullets]
                    tf_desc.TextRange.Text = "\r".join(all_lines)

                    for pi in range(1, tf_desc.TextRange.Paragraphs().Count + 1):
                        p = tf_desc.TextRange.Paragraphs(pi)
                        p.Font.Name = FONT_NAME
                        p.Font.Size = 11
                        p.Font.Color.RGB = dark_bgr
                        if pi == 1:
                            p.Font.Bold = True
                            p.ParagraphFormat.Bullet.Visible = False
                        else:
                            p.Font.Bold = False
                            p.ParagraphFormat.Bullet.Visible = False  # bullet char is in text
                        p.ParagraphFormat.SpaceBefore = 0
                        p.ParagraphFormat.SpaceAfter = 0
                else:
                    cell_date.TextFrame.TextRange.Text = ""
                    cell_desc.TextFrame.TextRange.Text = ""

        # Save and close presentation
        pres.Save()
        pres.Close()
        pres = None

        if created_app and ppt_app:
            try:
                ppt_app.Quit()
            except Exception:
                pass
            ppt_app = None

    except Exception as e:
        if pres is not None:
            try:
                pres.Close()
            except Exception:
                pass
        if created_app and ppt_app is not None:
            try:
                ppt_app.Quit()
            except Exception:
                pass
        raise e
    finally:
        pythoncom.CoUninitialize()

    return abs_output_path
