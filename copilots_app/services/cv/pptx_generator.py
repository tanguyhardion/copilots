"""
pptx_generator.py
Generates a 1-slide Executive Proposal CV in PowerPoint (.pptx) format
by precisely modifying template.pptx in-place to preserve 100% of the
original layout, geometry, positioning, typography, colors, and styling.
"""

from __future__ import annotations

import io
import os
import sys
from pathlib import Path
from typing import Any, Dict, List, Optional
import pptx
from pptx.dml.color import RGBColor
from pptx.util import Pt, Emu
from pptx.enum.text import PP_ALIGN
from lxml import etree

from copilots_app.services.cv.i18n import (
    PPTX_SECTION_DEFAULTS,
    get_cv_strings,
    normalize_sections,
)

FONT_NAME = "Calibri"
COLOR_DARK = RGBColor(36, 36, 36)  # RGB(36, 36, 36) matching template.pptx runs
COLOR_BLACK = RGBColor(0, 0, 0)
NO_STYLE_GUID = "{2D5ABB26-0587-4C30-8999-92F81FD0307C}"

# XML namespaces used in DrawingML
_NSMAP = {
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "p": "http://schemas.openxmlformats.org/presentationml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "pic": "http://schemas.openxmlformats.org/drawingml/2006/picture",
}


def _remove_shape_outline(shape) -> None:
    """Replace any existing <a:ln> on a shape's spPr with a no-line element."""
    ns_a = "http://schemas.openxmlformats.org/drawingml/2006/main"
    ns_p = "http://schemas.openxmlformats.org/presentationml/2006/main"

    sp_pr_el = shape._element.find(f"{{{ns_p}}}spPr")
    if sp_pr_el is None:
        return
    # Remove existing <a:ln> elements
    for ln_el in sp_pr_el.findall(f"{{{ns_a}}}ln"):
        sp_pr_el.remove(ln_el)
    # Insert <a:ln w="0"><a:noFill/></a:ln> to explicitly suppress the outline
    ln_no = etree.SubElement(sp_pr_el, f"{{{ns_a}}}ln")
    ln_no.set("w", "0")
    etree.SubElement(ln_no, f"{{{ns_a}}}noFill")



def _style_cell_borders_and_fill(
    cell,
    fill_hex: Optional[str] = None,
    mar_emu: int = 0,
) -> None:
    """
    Format a table cell's borders and fill in strict OpenXML schema order.

    ISO/IEC 29500 CT_TableCellProperties sequence:
      1. lnL
      2. lnR
      3. lnT
      4. lnB
      5. lnTlToBr
      6. lnBlToTr
      7. [cell3D]
      8. [EG_FillProperties: noFill / solidFill ...]
      9. [headers]
      10. [extLst]

    Incorrect child order causes PowerPoint to flag the presentation as corrupt!
    """
    ns_a = "http://schemas.openxmlformats.org/drawingml/2006/main"
    tc_pr = cell._tc.get_or_add_tcPr()
    tc_pr.clear()  # Strip existing child elements cleanly

    # 1. Borders in exact schema order: w=0, noFill
    for tag in ("lnL", "lnR", "lnT", "lnB", "lnTlToBr", "lnBlToTr"):
        ln = etree.SubElement(tc_pr, f"{{{ns_a}}}{tag}")
        ln.set("w", "0")
        etree.SubElement(ln, f"{{{ns_a}}}noFill")

    # 2. Fill properties (comes AFTER border lines in the schema)
    if fill_hex:
        solid = etree.SubElement(tc_pr, f"{{{ns_a}}}solidFill")
        clr = etree.SubElement(solid, f"{{{ns_a}}}srgbClr")
        clr.set("val", fill_hex)
    else:
        etree.SubElement(tc_pr, f"{{{ns_a}}}noFill")

    # 3. Cell margin attributes on tcPr
    m = str(mar_emu)
    tc_pr.set("marL", m)
    tc_pr.set("marR", m)
    tc_pr.set("marT", m)
    tc_pr.set("marB", m)


def _remove_table_cell_shading(table) -> None:
    """
    Strip all row/column styling flags and remove fill and borders on every cell.
    Keeps tableStyleId intact as defined in template ppt/tableStyles.xml to avoid repair errors.
    """
    ns_a = "http://schemas.openxmlformats.org/drawingml/2006/main"
    tbl_el = table._tbl
    tbl_pr = tbl_el.find(f"{{{ns_a}}}tblPr")
    if tbl_pr is not None:
        # Clear styling flags that trigger banded rows/columns or header borders
        for attr in ("bandRow", "bandCol", "firstRow", "lastRow", "firstCol", "lastCol"):
            tbl_pr.attrib.pop(attr, None)
        # Replace the style ID with the "no style" GUID
        for style_id in tbl_pr.findall(f"{{{ns_a}}}tableStyleId"):
            style_id.text = NO_STYLE_GUID
        # If there was no tableStyleId, insert one pointing to "no style"
        if not tbl_pr.findall(f"{{{ns_a}}}tableStyleId"):
            sid = etree.SubElement(tbl_pr, f"{{{ns_a}}}tableStyleId")
            sid.text = NO_STYLE_GUID
        # Ensure tblPr-level fill is noFill
        for fill_tag in ("solidFill", "gradFill", "pattFill", "blipFill"):
            for el in tbl_pr.findall(f"{{{ns_a}}}{fill_tag}"):
                tbl_pr.remove(el)
        if not tbl_pr.findall(f"{{{ns_a}}}noFill"):
            etree.SubElement(tbl_pr, f"{{{ns_a}}}noFill")

    # Explicitly noFill every cell and clear all per-cell borders
    for row in table.rows:
        for cell in row.cells:
            _style_cell_borders_and_fill(cell, fill_hex=None, mar_emu=0)


def _apply_table_gray_shading(table, rgb_hex: str = "E8E8E8", margin_emu: int = 45720) -> None:
    """Apply uniform solid gray fill, slight margins, and zero-width borders to every cell."""
    for row in table.rows:
        for cell in row.cells:
            _style_cell_borders_and_fill(cell, fill_hex=rgb_hex, mar_emu=margin_emu)


def _center_table_text(table) -> None:
    """Center-align all paragraphs in every cell of a table."""
    ns_a = "http://schemas.openxmlformats.org/drawingml/2006/main"
    for row in table.rows:
        for cell in row.cells:
            for para in cell.text_frame.paragraphs:
                p_el = para._p
                p_pr = p_el.find(f"{{{ns_a}}}pPr")
                if p_pr is None:
                    p_pr = etree.SubElement(p_el, f"{{{ns_a}}}pPr")
                    p_el.insert(0, p_pr)
                p_pr.set("algn", "ctr")


def _replace_rect_with_picture(slide, shape, image_path: str) -> None:
    """
    Remove a rectangle placeholder shape and insert a real picture in its place,
    at the same position and size, so the user only needs to do 'Replace Image'.
    """
    left = shape.left
    top = shape.top
    width = shape.width
    height = shape.height

    # Add the picture to the slide
    slide.shapes.add_picture(image_path, left, top, width, height)

    # Remove the old rectangle shape
    sp_el = shape._element
    sp_el.getparent().remove(sp_el)


def _force_shape_text_black(shape) -> None:
    """
    Force every run in a shape's text frame to explicit black (#000000).

    The template rectangles inherit white text from their <p:style>
    <a:fontRef val="lt1"> (Light 1 theme colour). Writing an explicit
    <a:solidFill><a:srgbClr val="000000"/> on the run overrides the
    style-level colour so the label is always visible on a white slide.
    """
    ns_a = "http://schemas.openxmlformats.org/drawingml/2006/main"
    if not shape.has_text_frame:
        return
    for para in shape.text_frame.paragraphs:
        for run in para.runs:
            rpr = run._r.find(f"{{{ns_a}}}rPr")
            if rpr is None:
                continue
            # Remove any existing fill on the run
            for fill_tag in ("solidFill", "gradFill", "pattFill", "blipFill", "noFill"):
                for el in rpr.findall(f"{{{ns_a}}}{fill_tag}"):
                    rpr.remove(el)
            # Insert <a:solidFill><a:srgbClr val="000000"/> as first child
            solid = etree.Element(f"{{{ns_a}}}solidFill")
            clr = etree.SubElement(solid, f"{{{ns_a}}}srgbClr")
            clr.set("val", "000000")
            rpr.insert(0, solid)


def _remove_table_borders(table) -> None:
    """
    Remove every visible border from a table: both the table-level tblBorder
    element and the per-cell lnL / lnR / lnT / lnB / lnTlToBr / lnBlToTr
    attributes in tcPr.
    """
    ns_a = "http://schemas.openxmlformats.org/drawingml/2006/main"
    tbl_el = table._tbl

    # --- Table-level borders ---
    tbl_pr = tbl_el.find(f"{{{ns_a}}}tblPr")
    if tbl_pr is not None:
        for tbl_border in tbl_pr.findall(f"{{{ns_a}}}tblBorder"):
            tbl_pr.remove(tbl_border)
        # Insert an explicit tblBorder that sets every edge to noFill
        tbl_border_el = etree.SubElement(tbl_pr, f"{{{ns_a}}}tblBorder")
        for edge in ("left", "right", "top", "bottom", "insideH", "insideV"):
            ln = etree.SubElement(tbl_border_el, f"{{{ns_a}}}ln")
            ln.set("w", "0")
            etree.SubElement(ln, f"{{{ns_a}}}noFill")

    # --- Per-cell borders ---
    for row in table.rows:
        for cell in row.cells:
            _style_cell_borders_and_fill(cell, fill_hex=None, mar_emu=0)


def _apply_paragraph_bullet(paragraph, char: str = "•", margin_left: str = "285750", indent: str = "-285750") -> None:
    """Turn a paragraph into a bullet point using standard DrawingML bullet settings."""
    ns_a = "http://schemas.openxmlformats.org/drawingml/2006/main"
    p_pr = paragraph._p.get_or_add_pPr()
    p_pr.set("marL", margin_left)
    p_pr.set("indent", indent)
    # Remove any existing bullet tags
    for child in list(p_pr):
        if child.tag.endswith("buChar") or child.tag.endswith("buNone") or child.tag.endswith("buAutoNum"):
            p_pr.remove(child)
    bu_char = etree.SubElement(p_pr, f"{{{ns_a}}}buChar")
    bu_char.set("char", char)


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

    # -------------------------------------------------------------------------
    # Style fixes applied to every generated slide regardless of content
    # -------------------------------------------------------------------------

    # Remove outline from "Text Placeholder 2" (overall info box with name etc.)
    for s in slide.shapes:
        if s.name == "Text Placeholder 2":
            _remove_shape_outline(s)

    # Remove outline + force black text on "Rectangle 3" (Profile Summary label)
    for s in slide.shapes:
        if s.name == "Rectangle 3":
            _remove_shape_outline(s)
            _force_shape_text_black(s)

    # Remove outline from "TextBox 7" (Profile Summary content)
    for s in slide.shapes:
        if s.name == "TextBox 7":
            _remove_shape_outline(s)

    # Remove outline + force black text on "Rectangle 57" (Relevant Experience label)
    for s in slide.shapes:
        if s.name == "Rectangle 57":
            _remove_shape_outline(s)
            _force_shape_text_black(s)

    # Remove outline on "Rectangle 49" (Subject Matter Expertise label)
    for s in slide.shapes:
        if s.name == "Rectangle 49":
            _remove_shape_outline(s)

    # Remove cell shading and all borders from Table 71 (Relevant Experience)
    t71_style = [s for s in slide.shapes if s.name == "Table 71" and s.has_table]
    if t71_style:
        _remove_table_cell_shading(t71_style[0].table)
        _remove_table_borders(t71_style[0].table)

    # Remove theme shading from Table 72 (Subject Matter Expertise) —
    # Remove theme shading and borders from Table 72 (Subject Matter Expertise) —
    # actual gray fill will be applied per-cell after content is written
    t72_style = [s for s in slide.shapes if s.name == "Table 72" and s.has_table]
    if t72_style:
        _remove_table_cell_shading(t72_style[0].table)
        _remove_table_borders(t72_style[0].table)

    # Replace "Picture 4" rectangle with a real picture placeholder image
    photo_placeholder_path = _resolve_photo_placeholder_path()
    pic4_shapes = [s for s in slide.shapes if s.name == "Picture 4"]
    if pic4_shapes and photo_placeholder_path:
        _replace_rect_with_picture(slide, pic4_shapes[0], photo_placeholder_path)

    # -------------------------------------------------------------------------
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

        # Apply gray shading (E8E8E8), slight cell margins, and centered text
        # 45720 EMU ≈ 0.5 pt ≈ ~0.64 mm — a barely-perceptible but present margin
        _apply_table_gray_shading(table_skills, rgb_hex="E8E8E8", margin_emu=45720)
        _center_table_text(table_skills)

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
                # Ensure the title/position is not a bullet point
                ns_a = "http://schemas.openxmlformats.org/drawingml/2006/main"
                p_pr0 = p0._p.get_or_add_pPr()
                p_pr0.attrib.pop("marL", None)
                p_pr0.attrib.pop("indent", None)
                for child in list(p_pr0):
                    if child.tag.endswith("buChar") or child.tag.endswith("buAutoNum") or child.tag.endswith("buNone"):
                        p_pr0.remove(child)
                etree.SubElement(p_pr0, f"{{{ns_a}}}buNone")

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
                    _apply_paragraph_bullet(p_b)
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
