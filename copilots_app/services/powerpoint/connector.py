"""
PowerPoint COM Automation Connector for rendering shapes in PowerPoint.
"""

import os
import re
import hashlib
import tempfile
import requests
from typing import List, Dict, Any, Optional, Callable

from copilots_app.services.powerpoint.constants import (
    SLIDE_WIDTH,
    SLIDE_HEIGHT,
    PPT_THEME_MAP,
    MSO_SHAPE_MAP,
    THEME_COLORS,
)
from copilots_app.services.powerpoint.parser import (
    download_icon,
    colorize_svg,
    get_svg_aspect_ratio,
    inject_svg_color,
    inline_svg_to_tempfile,
    hex_to_rgb,
    rgb_to_bgr_int,
    is_theme_color,
    sort_shapes_for_render,
)


def _svg_to_png(svg_path: str, width_px: int = 64, height_px: int = 64) -> Optional[str]:
    try:
        import cairosvg

        png_path = svg_path.replace(".svg", f"_{width_px}x{height_px}.png")
        cairosvg.svg2png(
            url=svg_path,
            write_to=png_path,
            output_width=width_px,
            output_height=height_px,
        )
        return png_path
    except Exception:
        return None


def utf16_len(text: str) -> int:
    return len(text.encode("utf-16-le")) // 2


def resolve_alignment(shape_def: Dict[str, Any]) -> Dict[str, Any]:
    align = shape_def.get("align", "")
    width = shape_def.get("width", 100)
    height = shape_def.get("height", 50)
    if not align:
        return shape_def
    if "center_x" in align or align == "center_xy":
        shape_def["left"] = (SLIDE_WIDTH - width) / 2
    if "center_y" in align or align == "center_xy":
        shape_def["top"] = (SLIDE_HEIGHT - height) / 2
    return shape_def


class PowerPointConnector:
    """Manages PowerPoint COM automation for creating slides and shapes."""

    def __init__(self):
        self.ppt_app = None

    def connect(self):
        import win32com.client
        import pythoncom

        pythoncom.CoInitialize()
        try:
            self.ppt_app = win32com.client.GetActiveObject("PowerPoint.Application")
        except Exception:
            self.ppt_app = win32com.client.Dispatch("PowerPoint.Application")
            self.ppt_app.Visible = True
        return self.ppt_app

    def _prefetch_icons(self, shapes: List[Dict[str, Any]], status_cb: Optional[Callable[[str], None]] = None) -> Dict[tuple, str]:
        icon_shapes = [s for s in shapes if s.get("type") == "icon"]
        if not icon_shapes:
            return {}
        cache = {}
        total = len(icon_shapes)
        for i, s in enumerate(icon_shapes):
            name = s.get("icon_name")
            style = s.get("icon_style", "solid")
            key = (name, style)
            if key not in cache:
                if status_cb:
                    status_cb(f"Downloading icon {i+1}/{total}: {name} ({style})…")
                cache[key] = download_icon(name, style)
        return cache

    def create_shapes_and_copy(self, shapes: List[Dict[str, Any]], status_cb: Optional[Callable[[str], None]] = None) -> bool:
        import pythoncom

        pythoncom.CoInitialize()
        try:
            icon_cache = self._prefetch_icons(shapes, status_cb)
            if status_cb:
                status_cb("Connecting to PowerPoint…")
            self.connect()
            if status_cb:
                status_cb("Creating temporary presentation…")
            temp_pres = self.ppt_app.Presentations.Add(WithWindow=False)
            temp_pres.PageSetup.SlideWidth = SLIDE_WIDTH
            temp_pres.PageSetup.SlideHeight = SLIDE_HEIGHT
            slide = temp_pres.Slides.Add(1, 12)

            ordered_shapes = sort_shapes_for_render(shapes)
            total = len(ordered_shapes)
            for i, sd in enumerate(ordered_shapes):
                if status_cb:
                    status_cb(f"Creating shape {i+1}/{total}…")
                self._create_single_shape(slide, sd, icon_cache)

            if status_cb:
                status_cb("Copying to clipboard…")
            slide.Shapes.Range().Copy()
            temp_pres.Close()
            if status_cb:
                status_cb(f"✓ {total} shapes copied — paste with Ctrl+V")
            return True
        except Exception as e:
            if status_cb:
                status_cb(f"Error: {e}")
            raise
        finally:
            pythoncom.CoUninitialize()

    def create_on_new_slide(self, slides_data: List[List[Dict[str, Any]]], status_cb: Optional[Callable[[str], None]] = None):
        import pythoncom

        pythoncom.CoInitialize()
        try:
            if slides_data and isinstance(slides_data[0], dict):
                slides_data = [slides_data]

            all_shapes = [s for slide in slides_data for s in slide]
            icon_cache = self._prefetch_icons(all_shapes, status_cb)

            self.connect()
            ppt = self.ppt_app
            if ppt.Presentations.Count == 0:
                raise Exception("No presentation open. Please open one in PowerPoint first.")
            pres = ppt.ActivePresentation

            try:
                current_index = ppt.ActiveWindow.View.Slide.SlideIndex
            except Exception:
                current_index = pres.Slides.Count

            total_slides = len(slides_data)

            for slide_num, shapes in enumerate(slides_data):
                new_index = current_index + slide_num + 1
                if status_cb:
                    status_cb(f"Creating slide {slide_num + 1}/{total_slides} (index {new_index})…")
                slide = pres.Slides.Add(new_index, 12)

                try:
                    ppt.ActiveWindow.View.GotoSlide(new_index)
                    ppt.ActiveWindow.Activate()
                except Exception:
                    pass

                ordered = sort_shapes_for_render(shapes)
                total_shapes = len(ordered)
                for i, sd in enumerate(ordered):
                    if status_cb:
                        status_cb(f"Slide {slide_num + 1}/{total_slides} — shape {i + 1}/{total_shapes}…")
                    self._create_single_shape(slide, sd, icon_cache)

            if status_cb:
                status_cb(f"✓ {total_slides} slide(s) created ({sum(len(s) for s in slides_data)} shapes total)!")
        finally:
            pythoncom.CoUninitialize()

    def create_on_current_slide(self, shapes: List[Dict[str, Any]], status_cb: Optional[Callable[[str], None]] = None):
        import pythoncom

        pythoncom.CoInitialize()
        try:
            icon_cache = self._prefetch_icons(shapes, status_cb)
            self.connect()
            ppt = self.ppt_app

            if ppt.Presentations.Count == 0:
                raise Exception("No presentation open. Please open one in PowerPoint first.")

            try:
                slide = ppt.ActiveWindow.View.Slide
                slide_index = slide.SlideIndex
            except Exception:
                raise Exception("Could not get active slide. Please click on a slide in PowerPoint first.")

            ordered_shapes = sort_shapes_for_render(shapes)
            total = len(ordered_shapes)
            for i, sd in enumerate(ordered_shapes):
                if status_cb:
                    status_cb(f"Inserting shape {i+1}/{total} on current slide…")
                self._create_single_shape(slide, sd, icon_cache)

            if status_cb:
                status_cb(f"✓ {total} shapes inserted on slide {slide_index}!")
        finally:
            pythoncom.CoUninitialize()

    def _create_single_shape(self, slide, shape_def: Dict[str, Any], icon_cache: Optional[Dict[tuple, str]] = None):
        shape_type = shape_def.get("type", "rect")

        if shape_type == "icon":
            return self._create_icon(slide, shape_def, icon_cache or {})
        if shape_type == "svg":
            return self._create_svg(slide, shape_def)
        if shape_type == "table":
            return self._create_table(slide, shape_def)
        if shape_type == "line":
            return self._create_line(slide, shape_def)
        elif shape_type == "image":
            return self._create_image(slide, shape_def)

        shape_def = resolve_alignment(dict(shape_def))
        left = shape_def.get("left", 0)
        top = shape_def.get("top", 0)
        width = shape_def.get("width", 100)
        height = shape_def.get("height", 50)
        color = shape_def.get("color") or "#CCCCCC"
        rotation = shape_def.get("rotation", 0)

        try:
            if shape_type == "text":
                ppt_shape = slide.Shapes.AddTextbox(1, left, top, width, height)
                ppt_shape.Fill.Visible = False
                ppt_shape.Line.Visible = False
            else:
                mso = MSO_SHAPE_MAP.get(shape_type, 1)
                ppt_shape = slide.Shapes.AddShape(mso, left, top, width, height)

                gradient = shape_def.get("gradient")
                if gradient:
                    self._apply_gradient(ppt_shape, gradient)
                else:
                    if is_theme_color(color):
                        ppt_shape.Fill.ForeColor.ObjectThemeColor = PPT_THEME_MAP[color]
                    else:
                        r, g, b = hex_to_rgb(color)
                        ppt_shape.Fill.ForeColor.RGB = rgb_to_bgr_int(r, g, b)
                    ppt_shape.Fill.Solid()
                    ppt_shape.Fill.Transparency = shape_def.get("transparency", 0.0)

                outline = shape_def.get("outline")
                if outline:
                    ppt_shape.Line.Visible = True
                    ppt_shape.Line.Weight = outline.get("weight", 1.0)
                    ol_color = outline.get("color") or "#333333"
                    if is_theme_color(ol_color):
                        ppt_shape.Line.ForeColor.ObjectThemeColor = PPT_THEME_MAP[ol_color]
                    else:
                        r, g, b = hex_to_rgb(ol_color)
                        ppt_shape.Line.ForeColor.RGB = rgb_to_bgr_int(r, g, b)
                else:
                    ppt_shape.Line.Visible = False

                if shape_type == "rounded_rect":
                    br = shape_def.get("border_radius")
                    if br is not None:
                        try:
                            min_dim = min(width, height)
                            if min_dim > 0:
                                val = float(min(0.5, br / min_dim))
                                ppt_shape.Adjustments.Item(1, val)
                        except Exception:
                            pass

                shadow = shape_def.get("shadow")
                if shadow:
                    self._apply_shadow(ppt_shape, shadow)

            if ppt_shape and rotation:
                ppt_shape.Rotation = rotation
            if ppt_shape and "rich_text" in shape_def:
                self._apply_rich_text(ppt_shape, shape_def)
            elif ppt_shape and "text" in shape_def:
                self._apply_simple_text(ppt_shape, shape_def)

            return ppt_shape
        except Exception as e:
            src = shape_def.get("_source_line", "?")
            print(f"[ppt] Error on line {src} ('{shape_type}'): {e}")
            return None

    def _create_icon(self, slide, shape_def, icon_cache):
        name = shape_def.get("icon_name", "question")
        style = shape_def.get("icon_style", "solid")
        left = shape_def.get("left", 0)
        top = shape_def.get("top", 0)
        width = shape_def.get("width", 48)
        color = shape_def.get("icon_color")

        svg_path = icon_cache.get((name, style))
        if not svg_path or not os.path.exists(svg_path):
            return self._icon_placeholder(slide, shape_def)

        working_path = colorize_svg(svg_path, color) if color else svg_path
        aspect_ratio = get_svg_aspect_ratio(working_path)
        if aspect_ratio <= 0:
            aspect_ratio = 1.0

        final_width = width
        final_height = width / aspect_ratio

        png_path = _svg_to_png(
            working_path,
            width_px=max(1, int(final_width * 2)),
            height_px=max(1, int(final_height * 2)),
        )
        insert_path = png_path if png_path and os.path.exists(png_path) else working_path
        try:
            pic = slide.Shapes.AddPicture(
                FileName=os.path.abspath(insert_path),
                LinkToFile=False,
                SaveWithDocument=True,
                Left=left,
                Top=top,
                Width=final_width,
                Height=final_height,
            )
            pic.LockAspectRatio = True
            return pic
        except Exception:
            return self._icon_placeholder(slide, shape_def)

    def _create_svg(self, slide, shape_def):
        left = shape_def.get("left", 0)
        top = shape_def.get("top", 0)
        width = shape_def.get("width", 48)
        height = shape_def.get("height", 48)
        rotation = shape_def.get("rotation", 0)
        svg_markup = shape_def.get("svg_markup", "")
        svg_color = shape_def.get("svg_color")

        if not svg_markup.strip():
            return None

        if svg_color:
            svg_markup = inject_svg_color(svg_markup, svg_color)

        svg_path = inline_svg_to_tempfile(svg_markup, shape_def.get("_source_line"))
        if not svg_path or not os.path.exists(svg_path):
            return None

        fit = shape_def.get("fit", "contain")
        aspect_ratio = get_svg_aspect_ratio(svg_path)
        if aspect_ratio <= 0:
            aspect_ratio = 1.0

        render_width, render_height = width, height
        if fit == "contain":
            if width / aspect_ratio <= height:
                render_height = width / aspect_ratio
            else:
                render_width = height * aspect_ratio
        elif fit == "cover":
            if width / aspect_ratio >= height:
                render_height = width / aspect_ratio
            else:
                render_width = height * aspect_ratio

        try:
            pic = slide.Shapes.AddPicture(
                FileName=os.path.abspath(svg_path),
                LinkToFile=False,
                SaveWithDocument=True,
                Left=left,
                Top=top,
                Width=render_width,
                Height=render_height,
            )
            pic.LockAspectRatio = True
            if rotation:
                pic.Rotation = rotation
            return pic
        except Exception:
            png_path = _svg_to_png(svg_path, width_px=max(1, int(render_width * 2)), height_px=max(1, int(render_height * 2)))
            if png_path and os.path.exists(png_path):
                try:
                    pic = slide.Shapes.AddPicture(
                        FileName=os.path.abspath(png_path),
                        LinkToFile=False,
                        SaveWithDocument=True,
                        Left=left,
                        Top=top,
                        Width=render_width,
                        Height=render_height,
                    )
                    pic.LockAspectRatio = True
                    if rotation:
                        pic.Rotation = rotation
                    return pic
                except Exception:
                    pass
            return None

    def _icon_placeholder(self, slide, shape_def):
        left = shape_def.get("left", 0)
        top = shape_def.get("top", 0)
        width = shape_def.get("width", 48)
        height = shape_def.get("height", 48)
        name = shape_def.get("icon_name", "?")

        ph = slide.Shapes.AddShape(1, left, top, width, height)
        ph.Fill.ForeColor.RGB = rgb_to_bgr_int(200, 200, 200)
        ph.Fill.Solid()
        ph.Line.Visible = True
        ph.Line.Weight = 1.0
        r, g, b = hex_to_rgb("#999999")
        ph.Line.ForeColor.RGB = rgb_to_bgr_int(r, g, b)
        tf = ph.TextFrame
        tf.TextRange.Text = name
        tf.TextRange.Font.Size = max(7, int(min(width, height) * 0.18))
        tf.TextRange.Font.Bold = False
        tf.VerticalAnchor = 3
        for pi in range(1, tf.TextRange.Paragraphs().Count + 1):
            tf.TextRange.Paragraphs(pi).ParagraphFormat.Alignment = 2
        return ph

    def _apply_shadow(self, ppt_shape, shadow):
        try:
            s = ppt_shape.Shadow
            s.Visible = True
            s.OffsetX = shadow.get("offset_x", 3)
            s.OffsetY = shadow.get("offset_y", 3)
            s.Blur = shadow.get("blur", 4)
            sc = shadow.get("color", "#333333")
            r, g, b = hex_to_rgb(sc[:7] if sc.startswith("#") else sc)
            s.ForeColor.RGB = rgb_to_bgr_int(r, g, b)
        except Exception:
            pass

    def _apply_gradient(self, ppt_shape, gradient):
        try:
            raw_stops = gradient.get("stops", ["#FFFFFF", "#000000"])
            stops = []
            for s in raw_stops:
                if not s:
                    continue
                if is_theme_color(s):
                    base_hex = THEME_COLORS.get(s)
                    if base_hex and re.match(r"^#[0-9A-Fa-f]{6}$", base_hex):
                        stops.append(base_hex)
                elif re.match(r"^#[0-9A-Fa-f]{6}$", s):
                    stops.append(s)
            if not stops:
                return

            angle = float(gradient.get("angle", 90))
            n = len(stops)

            if n < 2:
                r, g, b = hex_to_rgb(stops[0])
                ppt_shape.Fill.ForeColor.RGB = rgb_to_bgr_int(r, g, b)
                ppt_shape.Fill.Solid()
                return

            r0, g0, b0 = hex_to_rgb(stops[0])
            ppt_shape.Fill.ForeColor.RGB = rgb_to_bgr_int(r0, g0, b0)
            ppt_shape.Fill.OneColorGradient(1, 1, 1.0)
            try:
                ppt_shape.Fill.GradientAngle = angle
            except Exception:
                pass

            gs = ppt_shape.Fill.GradientStops
            while gs.Count > 2:
                gs.Delete(gs.Count)

            for idx, hex_color in enumerate(stops):
                position = max(0.01, min(0.99, idx / (n - 1)))
                r, g, b = hex_to_rgb(hex_color)
                rgb_val = rgb_to_bgr_int(r, g, b)
                if idx < gs.Count:
                    gs(idx + 1).Color.RGB = rgb_val
                    gs(idx + 1).Position = position
                else:
                    gs.Insert(rgb_val, position)

            while gs.Count > n:
                gs.Delete(gs.Count)
        except Exception:
            pass

    def _create_table(self, slide, shape_def):
        left = shape_def.get("left", 0)
        top = shape_def.get("top", 0)
        width = shape_def.get("width", 400)
        height = shape_def.get("height", 200)
        td = shape_def.get("table", {})
        style = shape_def.get("table_style", {})
        rows = td.get("rows", 2)
        cols = td.get("cols", 2)

        ppt_shape = slide.Shapes.AddTable(rows, cols, left, top, width, height)
        table = ppt_shape.Table

        col_widths = td.get("col_widths")
        if col_widths:
            total_w = sum(col_widths)
            if total_w > 0:
                scale = width / total_w
                for ci, cw in enumerate(col_widths):
                    if ci < cols:
                        table.Columns(ci + 1).Width = cw * scale

        content = td.get("content", [])
        for ri in range(min(rows, len(content))):
            for ci in range(min(cols, len(content[ri]))):
                cell = table.Cell(ri + 1, ci + 1)
                cell_shape = cell.Shape
                cell_shape.TextFrame.TextRange.Text = str(content[ri][ci])

                is_header = ri == 0 and td.get("header_row", False)
                fill_color = None
                if style:
                    if is_header:
                        fill_color = style.get("header_fill")
                    elif ri % 2 == 0:
                        fill_color = style.get("row_fill")
                    else:
                        fill_color = style.get("alt_row_fill") or style.get("row_fill")

                if fill_color:
                    try:
                        if is_theme_color(fill_color):
                            cell_shape.Fill.ForeColor.ObjectThemeColor = PPT_THEME_MAP[fill_color]
                        else:
                            r, g, b = hex_to_rgb(fill_color)
                            cell_shape.Fill.ForeColor.RGB = rgb_to_bgr_int(r, g, b)
                        cell_shape.Fill.Solid()
                    except Exception:
                        pass

                if style:
                    try:
                        border_color = style.get("border_color")
                        border_weight = style.get("border_weight", 1.0)
                        if border_color:
                            if is_theme_color(border_color):
                                cell_shape.Line.ForeColor.ObjectThemeColor = PPT_THEME_MAP[border_color]
                            else:
                                r, g, b = hex_to_rgb(border_color)
                                cell_shape.Line.ForeColor.RGB = rgb_to_bgr_int(r, g, b)
                            cell_shape.Line.Visible = True
                            cell_shape.Line.Weight = border_weight
                    except Exception:
                        pass

                    try:
                        font = cell_shape.TextFrame.TextRange.Font
                        if style.get("font"):
                            font.Name = style["font"]
                        if style.get("font_size"):
                            font.Size = style["font_size"]
                        if is_header:
                            font.Bold = style.get("header_bold", True)

                        text_color = style.get("header_text_color") if is_header and style.get("header_text_color") else style.get("text_color")
                        if text_color:
                            if is_theme_color(text_color):
                                font.Color.ObjectThemeColor = PPT_THEME_MAP[text_color]
                            else:
                                r, g, b = hex_to_rgb(text_color)
                                font.Color.RGB = rgb_to_bgr_int(r, g, b)
                    except Exception:
                        pass

                    try:
                        cell_shape.TextFrame.VerticalAnchor = 3
                        cell_shape.TextFrame.TextRange.ParagraphFormat.Alignment = 2
                    except Exception:
                        pass
        return ppt_shape

    def _create_line(self, slide, shape_def):
        x1 = shape_def.get("left", 0)
        y1 = shape_def.get("top", 0)
        x2 = shape_def.get("x2", x1 + 100)
        y2 = shape_def.get("y2", y1)

        ppt_shape = slide.Shapes.AddLine(x1, y1, x2, y2)
        lc = shape_def.get("color") or "#333333"
        if is_theme_color(lc):
            ppt_shape.Line.ForeColor.ObjectThemeColor = PPT_THEME_MAP[lc]
        else:
            r, g, b = hex_to_rgb(lc)
            ppt_shape.Line.ForeColor.RGB = rgb_to_bgr_int(r, g, b)
        ppt_shape.Line.Weight = shape_def.get("line_weight", 1.5)
        dash_map = {"solid": 1, "dash": 4, "dot": 3, "dash_dot": 5}
        ppt_shape.Line.DashStyle = dash_map.get(shape_def.get("dash_style", "solid"), 1)
        return ppt_shape

    def _create_image(self, slide, shape_def):
        url = shape_def.get("url")
        try:
            headers = {
                "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36"
            }
            resp = requests.get(url, timeout=15, verify=False, headers=headers)
            resp.raise_for_status()

            from urllib.parse import urlparse
            parsed_url = urlparse(url)
            filename = os.path.basename(parsed_url.path)
            filename = re.sub(r'[<>:"/\\|?*]', "_", filename)
            if not filename:
                filename = f"image_{hashlib.md5(url.encode()).hexdigest()[:8]}.img"

            temp_path = os.path.join(tempfile.gettempdir(), filename)
            with open(temp_path, "wb") as f:
                f.write(resp.content)

            desired_left = float(shape_def.get("left", 0))
            desired_top = float(shape_def.get("top", 0))
            desired_width = float(shape_def.get("width", 100))
            desired_height = float(shape_def.get("height", 100))
            rotation = float(shape_def.get("rotation", 0))

            pic = slide.Shapes.AddPicture(
                FileName=os.path.abspath(temp_path),
                LinkToFile=False,
                SaveWithDocument=True,
                Left=desired_left,
                Top=desired_top,
                Width=-1,
                Height=-1,
            )
            pic.LockAspectRatio = True

            nat_w = float(pic.Width)
            nat_h = float(pic.Height)

            if nat_w > 0 and nat_h > 0:
                scale = min(desired_width / nat_w, desired_height / nat_h)
                final_w = nat_w * scale
                final_h = nat_h * scale
                pic.LockAspectRatio = False
                pic.Width = final_w
                pic.Height = final_h
                pic.LockAspectRatio = True
                pic.Left = desired_left + (desired_width - final_w) / 2
                pic.Top = desired_top + (desired_height - final_h) / 2

            if rotation:
                pic.Rotation = rotation

            return pic
        except Exception as e:
            print(f"[image] Failed to download {url}: {e}")
            return None

    def _apply_rich_text(self, ppt_shape, shape_def):
        rich_text = shape_def.get("rich_text", [])
        shape_type = shape_def.get("type", "rect")
        default_font = shape_def.get("font")

        tf = ppt_shape.TextFrame
        tf.WordWrap = True
        tf.AutoSize = 1 if shape_type == "text" else 0
        tf.MarginLeft = shape_def.get("padding_left", 0)
        tf.MarginRight = shape_def.get("padding_right", 0)
        tf.MarginTop = shape_def.get("padding_top", 0)
        tf.MarginBottom = shape_def.get("padding_bottom", 0)

        default_v = "top" if shape_type == "text" else "middle"
        v_map = {"top": 1, "middle": 3, "bottom": 4}
        tf.VerticalAnchor = v_map.get(shape_def.get("vertical_align", default_v), v_map[default_v])

        tf.TextRange.Text = ""

        for idx, seg in enumerate(rich_text):
            text = seg.get("text", "")
            if not text:
                continue

            if idx == 0:
                tf.TextRange.Text = text
                tr = tf.TextRange
            else:
                tr = tf.TextRange.InsertAfter(text)

            if "size" in seg:
                tr.Font.Size = seg["size"]
            if "bold" in seg:
                tr.Font.Bold = seg["bold"]
            if "italic" in seg:
                tr.Font.Italic = seg["italic"]
            if "underline" in seg:
                tr.Font.Underline = seg["underline"]
            if "font" in seg:
                tr.Font.Name = seg["font"]
            elif default_font:
                tr.Font.Name = default_font
            if "color" in seg:
                sc = seg["color"]
                if is_theme_color(sc):
                    tr.Font.Color.ObjectThemeColor = PPT_THEME_MAP[sc]
                else:
                    r, g, b = hex_to_rgb(sc)
                    tr.Font.Color.RGB = rgb_to_bgr_int(r, g, b)

        default_h = "left" if shape_type == "text" else "center"
        h_map = {"left": 1, "center": 2, "right": 3, "justify": 4}
        align_val = h_map.get(shape_def.get("text_align", default_h), h_map[default_h])
        bullet = shape_def.get("bullet", False)
        lh = shape_def.get("line_height")

        for pi in range(1, tf.TextRange.Paragraphs().Count + 1):
            para = tf.TextRange.Paragraphs(pi)
            para.ParagraphFormat.Alignment = align_val
            try:
                if bullet:
                    para.ParagraphFormat.Bullet.Visible = True
                    if isinstance(bullet, str) and bullet:
                        para.ParagraphFormat.Bullet.Character = ord(bullet[0])
                else:
                    para.ParagraphFormat.Bullet.Visible = False
            except Exception:
                pass
            if lh is not None:
                try:
                    lh_val = float(lh)
                    if lh_val > 0:
                        para.ParagraphFormat.SpaceWithin = lh_val
                except Exception:
                    pass

    def _apply_simple_text(self, ppt_shape, shape_def):
        shape_type = shape_def.get("type", "rect")
        tf = ppt_shape.TextFrame
        default_font = shape_def.get("font")
        tf.WordWrap = True
        tf.AutoSize = 1 if shape_type == "text" else 0
        tf.MarginLeft = shape_def.get("padding_left", 0)
        tf.MarginRight = shape_def.get("padding_right", 0)
        tf.MarginTop = shape_def.get("padding_top", 0)
        tf.MarginBottom = shape_def.get("padding_bottom", 0)
        tf.TextRange.Text = shape_def.get("text", "")

        if default_font:
            tf.TextRange.Font.Name = default_font
        if "font_size" in shape_def:
            tf.TextRange.Font.Size = shape_def["font_size"]

        default_v = "top" if shape_type == "text" else "middle"
        v_map = {"top": 1, "middle": 3, "bottom": 4}
        tf.VerticalAnchor = v_map.get(shape_def.get("vertical_align", default_v), v_map[default_v])

        default_h = "left" if shape_type == "text" else "center"
        h_map = {"left": 1, "center": 2, "right": 3, "justify": 4}
        align_val = h_map.get(shape_def.get("text_align", default_h), h_map[default_h])
        bullet = shape_def.get("bullet", False)
        lh = shape_def.get("line_height")

        for pi in range(1, tf.TextRange.Paragraphs().Count + 1):
            para = tf.TextRange.Paragraphs(pi)
            para.ParagraphFormat.Alignment = align_val
            try:
                if bullet:
                    para.ParagraphFormat.Bullet.Visible = True
                    if isinstance(bullet, str) and bullet:
                        para.ParagraphFormat.Bullet.Character = ord(bullet[0])
                else:
                    para.ParagraphFormat.Bullet.Visible = False
            except Exception:
                pass
            if lh is not None:
                try:
                    lh_val = float(lh)
                    if lh_val > 0:
                        para.ParagraphFormat.SpaceWithin = lh_val
                except Exception:
                    pass
