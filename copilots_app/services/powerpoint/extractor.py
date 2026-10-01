"""
PowerPoint Extractor: Inspects the active PowerPoint slide and extracts
its shapes, background, text frames, and formatting into clean DSL lines annotated with `// id=N`.
"""

import re
from typing import Dict, Any, List, Optional, Tuple

from copilots_app.services.powerpoint.constants import (
    SLIDE_WIDTH,
    SLIDE_HEIGHT,
    REVERSE_MSO_SHAPE_MAP,
    REVERSE_PPT_THEME_MAP,
    REVERSE_DSL_COLOR_ALIASES,
    THEME_COLORS,
)
from copilots_app.services.powerpoint.parser import bgr_int_to_hex
from copilots_app.services.powerpoint.animations import (
    extract_slide_animations,
    format_animation_dsl,
)


def color_int_to_dsl(fore_color) -> Optional[str]:
    """Convert a COM ForeColor object to either a DSL theme token or hex color string."""
    if fore_color is None:
        return None
    try:
        theme_idx = getattr(fore_color, "ObjectThemeColor", 0)
        if theme_idx and theme_idx in REVERSE_PPT_THEME_MAP:
            theme_key = REVERSE_PPT_THEME_MAP[theme_idx]
            alias = REVERSE_DSL_COLOR_ALIASES.get(theme_key, theme_key)
            return alias
    except Exception:
        pass

    try:
        rgb_int = fore_color.RGB
        hex_str = bgr_int_to_hex(rgb_int)
        if hex_str:
            # Check if this hex closely matches any known theme color
            for token, t_hex in THEME_COLORS.items():
                if t_hex.upper() == hex_str.upper():
                    return REVERSE_DSL_COLOR_ALIASES.get(token, hex_str)
            return hex_str
    except Exception:
        pass
    return None


class PowerPointExtractor:
    """Extracts shapes and content from the currently active slide in PowerPoint into DSL."""

    def __init__(self):
        self.ppt_app = None

    def connect(self):
        import win32com.client
        import pythoncom

        pythoncom.CoInitialize()
        try:
            self.ppt_app = win32com.client.GetActiveObject("PowerPoint.Application")
        except Exception:
            raise Exception("PowerPoint is not running. Please open PowerPoint and a presentation first.")
        return self.ppt_app

    def extract_active_slide(self) -> Tuple[str, Dict[str, Any]]:
        """
        Inspects the active slide and produces DSL lines annotated with `// id=<ShapeId>`.
        Returns (dsl_text, metadata_dict).
        """
        import pythoncom

        pythoncom.CoInitialize()
        try:
            self.connect()
            ppt = self.ppt_app
            if ppt.Presentations.Count == 0:
                raise Exception("No presentations are currently open in PowerPoint.")

            pres = ppt.ActivePresentation
            try:
                slide = ppt.ActiveWindow.View.Slide
                slide_index = slide.SlideIndex
            except Exception:
                raise Exception("Could not access active slide. Please select a slide in PowerPoint first.")

            slide_count = pres.Slides.Count
            dsl_text, items = self._extract_slide_dsl(
                slide, f"// === Extracted from Active Slide ({slide_index}/{slide_count}) ==="
            )
            meta = {
                "slide_index": slide_index,
                "slide_count": slide_count,
                "shape_count": len(items),
                "shapes": items,
            }
            return dsl_text, meta
        finally:
            pythoncom.CoUninitialize()

    def extract_all_slides(self) -> Tuple[str, Dict[str, Any]]:
        """
        Inspects every slide of the active presentation and produces DSL blocks separated
        by `---` (the multi-slide separator understood by the parser).
        Returns (dsl_text, metadata_dict).
        """
        import pythoncom

        pythoncom.CoInitialize()
        try:
            self.connect()
            ppt = self.ppt_app
            if ppt.Presentations.Count == 0:
                raise Exception("No presentations are currently open in PowerPoint.")

            pres = ppt.ActivePresentation
            slide_count = pres.Slides.Count
            blocks = []
            slides_meta = []
            for idx in range(1, slide_count + 1):
                slide = pres.Slides(idx)
                dsl_text, items = self._extract_slide_dsl(
                    slide, f"// === Slide {idx}/{slide_count} ==="
                )
                blocks.append(dsl_text)
                slides_meta.append({"slide_index": idx, "shape_count": len(items), "shapes": items})

            meta = {
                "slide_count": slide_count,
                "shape_count": sum(m["shape_count"] for m in slides_meta),
                "slides": slides_meta,
            }
            return "\n\n---\n\n".join(blocks), meta
        finally:
            pythoncom.CoUninitialize()

    def _extract_slide_dsl(self, slide, header: str) -> Tuple[str, List[Dict[str, Any]]]:
        """Produce the DSL text for a single slide object, plus the extracted shape metadata."""
        lines = [
            header,
            "// Edit mode: modify or replace shapes using id=N",
            "",
        ]

        # 1. Slide Background
        bg_dsl = self._extract_slide_background(slide)
        if bg_dsl:
            lines.append(bg_dsl)
            lines.append("")

        # 2. Extract Shapes and Animations
        shapes_count = slide.Shapes.Count
        slide_animations = extract_slide_animations(slide)
        extracted_items = []

        for i in range(1, shapes_count + 1):
            try:
                shape = slide.Shapes(i)
                anim_cfg = slide_animations.get(shape.Id)
                shape_dsl = self._extract_shape(shape, anim_cfg)
                if shape_dsl:
                    lines.append(shape_dsl)
                    extracted_items.append({
                        "id": shape.Id,
                        "name": shape.Name,
                        "type": getattr(shape, "Type", 0),
                        "animation": anim_cfg,
                    })
            except Exception as e:
                print(f"[extractor] Error extracting shape index {i}: {e}")

        return "\n".join(lines).strip(), extracted_items

    def _extract_slide_background(self, slide) -> Optional[str]:
        try:
            follow = getattr(slide, "FollowMasterBackground", True)
            if not follow:
                fill = slide.Background.Fill
                color_token = color_int_to_dsl(fill.ForeColor)
                if color_token:
                    return f"slide background={color_token}  // id=slide"
        except Exception as e:
            print(f"[extractor] Error extracting slide background: {e}")
        return None

    def _extract_shape(self, shape, anim_cfg: Optional[Dict[str, Any]] = None) -> Optional[str]:
        shape_id = shape.Id
        shape_type_int = getattr(shape, "Type", 1)

        # Tables (type 19 or HasTable)
        has_table = False
        try:
            has_table = shape.HasTable == -1 or shape_type_int == 19
        except Exception:
            pass
        if has_table:
            return self._extract_table(shape, shape_id, anim_cfg)

        # Lines (type 9)
        if shape_type_int == 9:
            return self._extract_line(shape, shape_id, anim_cfg)

        # Images (type 11 or 13)
        if shape_type_int in (11, 13):
            return self._extract_image(shape, shape_id, anim_cfg)

        # Standard shapes / textboxes / placeholders
        return self._extract_standard_shape(shape, shape_id, shape_type_int, anim_cfg)

    def _extract_standard_shape(self, shape, shape_id: int, shape_type_int: int, anim_cfg: Optional[Dict[str, Any]] = None) -> str:
        # Determine shape type keyword
        shape_type = "rect"
        if shape_type_int == 17:  # msoTextBox
            shape_type = "text"
        elif shape_type_int in (1, 14):  # msoAutoShape or msoPlaceholder
            auto_type = getattr(shape, "AutoShapeType", 1)
            # Placeholder without fill often acts like a text box
            try:
                fill_visible = shape.Fill.Visible != 0
            except Exception:
                fill_visible = True
            if shape_type_int == 14 and not fill_visible:
                shape_type = "text"
            else:
                shape_type = REVERSE_MSO_SHAPE_MAP.get(auto_type, "rect")

        # Geometry
        left = round(float(shape.Left), 1)
        top = round(float(shape.Top), 1)
        width = round(float(shape.Width), 1)
        height = round(float(shape.Height), 1)

        tokens = [shape_type, f"left={left}", f"top={top}", f"width={width}", f"height={height}"]

        try:
            rot = round(float(shape.Rotation), 1)
            if rot != 0:
                tokens.append(f"rotation={rot}")
        except Exception:
            pass

        # Animation
        if anim_cfg:
            anim_str = format_animation_dsl(anim_cfg)
            if anim_str:
                tokens.append(anim_str)

        # Fill color (if not plain text box or if fill is visible)
        try:
            if shape_type != "text" and shape.Fill.Visible != 0:
                color_token = color_int_to_dsl(shape.Fill.ForeColor)
                if color_token:
                    tokens.append(f"color={color_token}")
                transparency = round(float(shape.Fill.Transparency), 2)
                if transparency > 0:
                    tokens.append(f"transparency={transparency}")
        except Exception:
            pass

        # Outline
        try:
            if shape_type != "text" and shape.Line.Visible != 0:
                line_color = color_int_to_dsl(shape.Line.ForeColor)
                weight = round(float(shape.Line.Weight), 1)
                if line_color:
                    if weight == 1.0:
                        tokens.append(f"outline={line_color}")
                    else:
                        tokens.append(f"outline={line_color},{weight}")
        except Exception:
            pass

        # Corner radius for rounded_rect
        if shape_type == "rounded_rect":
            try:
                adj = float(shape.Adjustments.Item(1))
                min_dim = min(width, height)
                br = round(adj * min_dim, 1)
                if br > 0:
                    tokens.append(f"border_radius={br}")
            except Exception:
                pass

        # Shadow
        try:
            if shape.Shadow.Visible != 0:
                ox = round(float(shape.Shadow.OffsetX), 1)
                oy = round(float(shape.Shadow.OffsetY), 1)
                blur = round(float(shape.Shadow.Blur), 1)
                sc = color_int_to_dsl(shape.Shadow.ForeColor) or "#333333"
                if ox == 3 and oy == 3 and blur == 4:
                    tokens.append("shadow=true")
                else:
                    tokens.append(f"shadow={ox},{oy},{blur},{sc}")
        except Exception:
            pass

        # Text Frame
        text_clause = ""
        try:
            if shape.HasTextFrame and shape.TextFrame.HasText:
                tf = shape.TextFrame
                # Alignment
                tr = tf.TextRange
                try:
                    align_val = tr.ParagraphFormat.Alignment
                    h_rev = {1: "left", 2: "center", 3: "right", 4: "justify"}
                    if align_val in h_rev:
                        default_h = "left" if shape_type == "text" else "center"
                        if h_rev[align_val] != default_h:
                            tokens.append(f"halign={h_rev[align_val]}")
                except Exception:
                    pass

                try:
                    v_val = tf.VerticalAnchor
                    v_rev = {1: "top", 3: "middle", 4: "bottom"}
                    if v_val in v_rev:
                        default_v = "top" if shape_type == "text" else "middle"
                        if v_rev[v_val] != default_v:
                            tokens.append(f"valign={v_rev[v_val]}")
                except Exception:
                    pass

                # Margins / Padding
                try:
                    ml = round(float(tf.MarginLeft), 1)
                    mr = round(float(tf.MarginRight), 1)
                    mt = round(float(tf.MarginTop), 1)
                    mb = round(float(tf.MarginBottom), 1)
                    if ml > 0 or mr > 0 or mt > 0 or mb > 0:
                        if ml == mr == mt == mb:
                            tokens.append(f"padding={ml}")
                        else:
                            tokens.append(f"padding={ml},{mr},{mt},{mb}")
                except Exception:
                    pass

                text_clause = self._extract_rich_text(tr)
        except Exception as e:
            print(f"[extractor] Text extract error: {e}")

        line = " ".join(tokens)
        if text_clause:
            line += f" | {text_clause}"
        line += f"  // id={shape_id}"
        return line

    def _extract_rich_text(self, text_range) -> str:
        try:
            runs_count = text_range.Runs().Count
        except Exception:
            runs_count = 0

        if runs_count <= 1:
            full_text = text_range.Text.replace('"', '\\"').replace("\r\n", "\\n").replace("\n", "\\n").replace("\r", "\\n")
            if not full_text:
                return ""
            attrs = []
            try:
                f_size = round(float(text_range.Font.Size))
                if f_size:
                    attrs.append(f"size={f_size}")
            except Exception:
                pass
            try:
                if text_range.Font.Bold != 0:
                    attrs.append("bold=true")
                if text_range.Font.Italic != 0:
                    attrs.append("italic=true")
            except Exception:
                pass
            try:
                tc = color_int_to_dsl(text_range.Font.Color)
                if tc:
                    attrs.append(f"color={tc}")
            except Exception:
                pass
            if attrs:
                return f'"{full_text}" ' + " ".join(attrs)
            return f'"{full_text}"'

        # Multi-run rich text
        segments = []
        for ri in range(1, runs_count + 1):
            try:
                run = text_range.Runs(ri)
                raw_text = run.Text.replace('"', '\\"').replace("\r\n", "\\n").replace("\n", "\\n").replace("\r", "\\n")
                if not raw_text:
                    continue
                attrs = []
                try:
                    f_size = round(float(run.Font.Size))
                    if f_size:
                        attrs.append(f"size={f_size}")
                except Exception:
                    pass
                try:
                    if run.Font.Bold != 0:
                        attrs.append("bold=true")
                    if run.Font.Italic != 0:
                        attrs.append("italic=true")
                except Exception:
                    pass
                try:
                    tc = color_int_to_dsl(run.Font.Color)
                    if tc:
                        attrs.append(f"color={tc}")
                except Exception:
                    pass

                seg = f'"{raw_text}"'
                if attrs:
                    seg += f" {' '.join(attrs)}"
                segments.append(seg)
            except Exception:
                pass

        if segments:
            return " + ".join(segments)

        # Fallback
        t = text_range.Text.replace('"', '\\"').replace("\n", "\\n")
        return f'"{t}"'

    def _extract_line(self, shape, shape_id: int, anim_cfg: Optional[Dict[str, Any]] = None) -> str:
        left = round(float(shape.Left), 1)
        top = round(float(shape.Top), 1)
        width = round(float(shape.Width), 1)
        height = round(float(shape.Height), 1)
        x2 = round(left + width, 1)
        y2 = round(top + height, 1)

        tokens = ["line", f"x1={left}", f"y1={top}", f"x2={x2}", f"y2={y2}"]
        if anim_cfg:
            anim_str = format_animation_dsl(anim_cfg)
            if anim_str:
                tokens.append(anim_str)
        try:
            lc = color_int_to_dsl(shape.Line.ForeColor)
            if lc:
                tokens.append(f"color={lc}")
            weight = round(float(shape.Line.Weight), 1)
            if weight != 1.5:
                tokens.append(f"weight={weight}")
        except Exception:
            pass

        return " ".join(tokens) + f"  // id={shape_id}"

    def _extract_image(self, shape, shape_id: int, anim_cfg: Optional[Dict[str, Any]] = None) -> str:
        left = round(float(shape.Left), 1)
        top = round(float(shape.Top), 1)
        width = round(float(shape.Width), 1)
        height = round(float(shape.Height), 1)
        anim_token = ""
        if anim_cfg:
            anim_str = format_animation_dsl(anim_cfg)
            if anim_str:
                anim_token = f" {anim_str}"
        return (
            f"// [EMBEDDED IMAGE — id={shape_id}]\n"
            f"image url=https://... left={left} top={top} width={width} height={height}{anim_token}  // id={shape_id}"
        )

    def _extract_table(self, shape, shape_id: int, anim_cfg: Optional[Dict[str, Any]] = None) -> str:
        table = shape.Table
        rows_cnt = table.Rows.Count
        cols_cnt = table.Columns.Count

        left = round(float(shape.Left), 1)
        top = round(float(shape.Top), 1)
        width = round(float(shape.Width), 1)
        height = round(float(shape.Height), 1)

        tokens = [f"table", f"left={left}", f"top={top}", f"width={width}", f"height={height}"]
        if anim_cfg:
            anim_str = format_animation_dsl(anim_cfg)
            if anim_str:
                tokens.append(anim_str)
        out = [" ".join(tokens) + f"  // id={shape_id}"]

        # Column widths
        col_widths = []
        for c in range(1, cols_cnt + 1):
            try:
                col_widths.append(str(round(float(table.Columns(c).Width), 1)))
            except Exception:
                col_widths.append("100")
        out.append("cols=" + ",".join(col_widths))

        # Rows and header
        for r in range(1, rows_cnt + 1):
            cell_texts = []
            for c in range(1, cols_cnt + 1):
                try:
                    cell_text = table.Cell(r, c).Shape.TextFrame.TextRange.Text
                    clean = cell_text.replace('"', '\\"').replace("\r\n", " ").replace("\n", " ").strip()
                    cell_texts.append(f'"{clean}"')
                except Exception:
                    cell_texts.append('""')

            prefix = "header=" if r == 1 else "row="
            out.append(prefix + ",".join(cell_texts))

        return "\n".join(out)

