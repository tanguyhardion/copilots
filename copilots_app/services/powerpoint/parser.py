"""
PowerPoint DSL Parser, color utilities, and geometry resolvers.
"""

import os
import re
import hashlib
import tempfile
import requests
import urllib3
from typing import List, Dict, Any, Optional, Tuple

from copilots_app.services.powerpoint.constants import (
    VALID_SHAPE_TYPES,
    DEFAULT_DSL_COLORS,
    DSL_COLOR_ALIASES,
    DASH_STYLES_DSL,
    SLIDE_WIDTH,
    SLIDE_HEIGHT,
    THEME_COLORS,
    PPT_THEME_MAP,
    MSO_SHAPE_MAP,
)

urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

# ── ICON CACHE ────────────────────────────────────────────────────────────────
ICON_CACHE_DIR = "icons"
FA_BASE_URL = (
    "https://raw.githubusercontent.com/FortAwesome/Font-Awesome/"
    "6.x/svgs/{style}/{icon}.svg"
)
VALID_ICON_STYLES = ["solid", "regular", "brands"]


def download_icon(icon_name: str, style: str = "solid") -> Optional[str]:
    if style not in VALID_ICON_STYLES:
        style = "solid"

    style_dir = os.path.join(ICON_CACHE_DIR, style)
    os.makedirs(style_dir, exist_ok=True)
    file_path = os.path.join(style_dir, f"{icon_name}.svg")

    if os.path.exists(file_path):
        return file_path

    url = FA_BASE_URL.format(style=style, icon=icon_name)
    try:
        response = requests.get(url, verify=False, timeout=10)
        if response.status_code == 200:
            with open(file_path, "wb") as f:
                f.write(response.content)
            return file_path
        else:
            for alt_style in VALID_ICON_STYLES:
                if alt_style == style:
                    continue
                alt_url = FA_BASE_URL.format(style=alt_style, icon=icon_name)
                r2 = requests.get(alt_url, verify=False, timeout=10)
                if r2.status_code == 200:
                    with open(file_path, "wb") as f:
                        f.write(r2.content)
                    return file_path
            return None
    except Exception as e:
        print(f"[icon] Download error for '{icon_name}': {e}")
        return None


def colorize_svg(svg_path: str, hex_color: str) -> str:
    try:
        with open(svg_path, "r", encoding="utf-8") as f:
            content = f.read()
        content = re.sub(r'fill="[^"]*"', f'fill="{hex_color}"', content)
        if "fill=" not in content:
            content = content.replace("<svg", f'<svg fill="{hex_color}"', 1)
        colored_path = svg_path.replace(".svg", f"_{hex_color.lstrip('#')}.svg")
        with open(colored_path, "w", encoding="utf-8") as f:
            f.write(content)
        return colored_path
    except Exception as e:
        print(f"[icon] Colorize warning: {e}")
        return svg_path


def get_svg_aspect_ratio(svg_path: str) -> float:
    try:
        with open(svg_path, "r", encoding="utf-8") as f:
            content = f.read()
        m = re.search(r'viewBox="[^"]*\s([\d.]+)\s+([\d.]+)"', content)
        if m:
            vb_w, vb_h = float(m.group(1)), float(m.group(2))
            if vb_w > 0 and vb_h > 0:
                return vb_w / vb_h
        mw = re.search(r'width="([\d.]+)(?:px)?"', content)
        mh = re.search(r'height="([\d.]+)(?:px)?"', content)
        if mw and mh:
            w, h = float(mw.group(1)), float(mh.group(1))
            if w > 0 and h > 0:
                return w / h
    except Exception as e:
        print(f"[icon] aspect-ratio warning: {e}")
    return 1.0


def inject_svg_color(svg_markup: str, hex_color: str) -> str:
    """Apply a best-effort color tint to inline SVG markup."""
    if not hex_color:
        return svg_markup

    try:
        color = hex_color.strip()
        if not color.startswith("#"):
            color = f"#{color}"

        if re.search(r"<svg\b[^>]*\bstyle=", svg_markup, flags=re.I):
            svg_markup = re.sub(
                r'(<svg\b[^>]*\bstyle=")(.*?)"',
                lambda m: f'{m.group(1)}{m.group(2)};color:{color}"',
                svg_markup,
                count=1,
                flags=re.I,
            )
        else:
            svg_markup = re.sub(
                r"<svg\b",
                f'<svg style="color:{color};"',
                svg_markup,
                count=1,
                flags=re.I,
            )

        svg_markup = svg_markup.replace('fill="currentColor"', f'fill="{color}"')
        svg_markup = svg_markup.replace('stroke="currentColor"', f'stroke="{color}"')
        return svg_markup
    except Exception as e:
        print(f"[svg] color injection warning: {e}")
        return svg_markup


def inline_svg_to_tempfile(svg_markup: str, source_line: Optional[int] = None) -> Optional[str]:
    try:
        cache_dir = os.path.join(tempfile.gettempdir(), "ppt_shape_generator_svg")
        os.makedirs(cache_dir, exist_ok=True)
        digest = hashlib.sha1(svg_markup.encode("utf-8")).hexdigest()
        file_path = os.path.join(cache_dir, f"inline_{digest}.svg")
        if not os.path.exists(file_path):
            with open(file_path, "w", encoding="utf-8") as f:
                f.write(svg_markup)
        return file_path
    except Exception as e:
        print(f"[svg] temp file warning on line {source_line}: {e}")
        return None


# ── COLOR RESOLUTION ─────────────────────────────────────────────────────────
def hex_to_rgb(hex_color: str) -> Tuple[int, int, int]:
    try:
        hex_color = re.sub(r"[^0-9A-Fa-f]", "", hex_color.strip().lstrip("#"))
        if len(hex_color) == 3:
            hex_color = "".join(c * 2 for c in hex_color)
        if len(hex_color) != 6:
            raise ValueError(f"bad hex length: '{hex_color}'")
        return tuple(int(hex_color[i : i + 2], 16) for i in (0, 2, 4))
    except Exception as e:
        return (128, 128, 128)


def clamp(v: float, lo: int = 0, hi: int = 255) -> int:
    return max(lo, min(hi, int(round(v))))


def rgb_to_hex(r: int, g: int, b: int) -> str:
    return f"#{clamp(r):02X}{clamp(g):02X}{clamp(b):02X}"


def rgb_to_bgr_int(r: int, g: int, b: int) -> int:
    return r + (g << 8) + (b << 16)



def lighten_hex(hex_color: str, amount: float) -> str:
    r, g, b = hex_to_rgb(hex_color)
    nr = r + (255 - r) * amount
    ng = g + (255 - g) * amount
    nb = b + (255 - b) * amount
    return rgb_to_hex(nr, ng, nb)


def darken_hex(hex_color: str, amount: float) -> str:
    r, g, b = hex_to_rgb(hex_color)
    nr = r * (1 - amount)
    ng = g * (1 - amount)
    nb = b * (1 - amount)
    return rgb_to_hex(nr, ng, nb)


def bgr_int_to_hex(rgb_int: int) -> Optional[str]:
    try:
        v = int(rgb_int)
        r = v & 0xFF
        g = (v >> 8) & 0xFF
        b = (v >> 16) & 0xFF
        return f"#{r:02X}{g:02X}{b:02X}"
    except Exception:
        return None


def dsl_resolve_color(c: Optional[str]) -> Optional[str]:
    if not c:
        return None

    c = c.strip().lower()

    if c in DSL_COLOR_ALIASES:
        return DSL_COLOR_ALIASES[c]

    m = re.match(r"^(a[1-6]|bg[12]|t[12])_(l1|l2|d1|d2)$", c)
    if m:
        base_alias, variant = m.groups()
        theme_token = DSL_COLOR_ALIASES.get(base_alias)
        if not theme_token:
            return None

        base_hex = THEME_COLORS.get(theme_token)
        if not base_hex:
            return None

        if variant == "l1":
            return lighten_hex(base_hex, 0.35)
        if variant == "l2":
            return lighten_hex(base_hex, 0.60)
        if variant == "d1":
            return darken_hex(base_hex, 0.25)
        if variant == "d2":
            return darken_hex(base_hex, 0.45)

    if re.match(r"^#[0-9a-f]{3}$", c):
        return "#" + "".join(ch * 2 for ch in c[1:])

    if re.match(r"^#[0-9a-f]{6}$", c):
        return c.upper()

    return None


def is_theme_color(c: str) -> bool:
    return c in PPT_THEME_MAP


def get_active_ppt_theme_colors() -> Dict[str, str]:
    colors = {
        "theme_bg1": DEFAULT_DSL_COLORS["bg1"],
        "theme_text1": DEFAULT_DSL_COLORS["t1"],
        "theme_bg2": DEFAULT_DSL_COLORS["bg2"],
        "theme_text2": DEFAULT_DSL_COLORS["t2"],
        "theme_accent1": DEFAULT_DSL_COLORS["a1"],
        "theme_accent2": DEFAULT_DSL_COLORS["a2"],
        "theme_accent3": DEFAULT_DSL_COLORS["a3"],
        "theme_accent4": DEFAULT_DSL_COLORS["a4"],
        "theme_accent5": DEFAULT_DSL_COLORS["a5"],
        "theme_accent6": DEFAULT_DSL_COLORS["a6"],
    }

    try:
        import pythoncom
        import win32com.client

        pythoncom.CoInitialize()
        try:
            ppt = win32com.client.GetActiveObject("PowerPoint.Application")
            if ppt.Presentations.Count > 0:
                pres = ppt.ActivePresentation
                scheme = pres.SlideMaster.Theme.ThemeColorScheme
                for theme_name, idx in PPT_THEME_MAP.items():
                    try:
                        office_color = scheme.Colors(idx)
                        hex_value = bgr_int_to_hex(office_color.RGB)
                        if hex_value:
                            colors[theme_name] = hex_value
                    except Exception:
                        pass
        finally:
            pythoncom.CoUninitialize()
    except Exception:
        pass

    return colors


def refresh_dsl_theme_colors():
    global THEME_COLORS
    THEME_COLORS = get_active_ppt_theme_colors()


# ── TOKENIZER & PARSER ────────────────────────────────────────────────────────
def _split_on_pipe(s: str) -> Tuple[str, Optional[str]]:
    in_quotes = False
    for i, ch in enumerate(s):
        if ch == '"':
            in_quotes = not in_quotes
        elif ch == "|" and not in_quotes:
            return s[:i], s[i + 1 :]
    return s, None


def _tokenize_kvs(s: str) -> List[str]:
    tokens = []
    pattern = re.compile(r'(\w[\w\-]*(?:=[^\s"]*|="[^"]*")?)')
    for m in pattern.finditer(s):
        tok = m.group(1).strip()
        if tok:
            tokens.append(tok)
    return tokens


def tokenize_dsl_line(line: str) -> Tuple[List[str], Optional[str]]:
    line = re.sub(r"(?<!\S)//.*$", "", line).strip()
    if not line:
        return [], None
    before, after = _split_on_pipe(line)
    tokens = _tokenize_kvs(before)
    return tokens, after.strip() if after is not None else None


def extract_fields(tokens: List[str]) -> Tuple[Optional[str], Dict[str, str]]:
    if not tokens:
        return None, {}
    shape_type = tokens[0].lower()
    fields = {}
    for tok in tokens[1:]:
        if "=" in tok:
            key, _, value = tok.partition("=")
            fields[key.strip().lower()] = value.strip().strip('"')
    return shape_type, fields


def parse_text_segment(seg: str) -> Optional[Dict[str, Any]]:
    seg = seg.strip()
    m = re.match(r'"((?:[^"\\]|\\.)*)"(.*)', seg)
    if not m:
        return None

    text = m.group(1).replace("\\n", "\n").replace('\\"', '"')
    rest = m.group(2).strip()
    result = {"text": text}

    if not rest:
        return result

    for tok in rest.split():
        if "=" not in tok:
            continue
        k, _, v = tok.partition("=")
        k, v = k.strip().lower(), v.strip()
        if k == "size":
            try:
                result["size"] = int(v)
            except ValueError:
                pass
        elif k == "bold":
            result["bold"] = v.lower() == "true"
        elif k == "italic":
            result["italic"] = v.lower() == "true"
        elif k == "underline":
            result["underline"] = v.lower() == "true"
        elif k == "color":
            resolved = dsl_resolve_color(v)
            if resolved:
                result["color"] = resolved
        elif k == "font":
            result["font"] = v

    return result


def parse_rich_text(text_part: str) -> List[Dict[str, Any]]:
    segments, current, in_quotes = [], "", False
    for ch in text_part:
        if ch == '"':
            in_quotes = not in_quotes
            current += ch
        elif ch == "+" and not in_quotes:
            segments.append(current.strip())
            current = ""
        else:
            current += ch
    if current.strip():
        segments.append(current.strip())

    return [r for r in (parse_text_segment(s) for s in segments) if r]


def build_shape_from_fields(shape_type: str, fields: Dict[str, str], text_part: Optional[str], source_line: Optional[int] = None) -> Dict[str, Any]:
    shape = {"type": shape_type}
    if source_line is not None:
        shape["_source_line"] = source_line

    for f in ("left", "top", "width", "height"):
        if f in fields:
            try:
                shape[f] = float(fields[f])
            except ValueError:
                pass

    if "color" in fields:
        resolved = dsl_resolve_color(fields["color"])
        if resolved:
            shape["color"] = resolved

    if "transparency" in fields:
        try:
            shape["transparency"] = float(fields["transparency"])
        except ValueError:
            pass

    if "grad_stops" in fields:
        raw_stops = fields["grad_stops"].split(",")
        stops = [dsl_resolve_color(c.strip()) for c in raw_stops]
        stops = [s for s in stops if s]
        if stops:
            angle = 90.0
            if "grad_angle" in fields:
                try:
                    angle = float(fields["grad_angle"])
                except ValueError:
                    pass
            shape["gradient"] = {"stops": stops, "angle": angle}

    if "outline" in fields:
        parts = fields["outline"].split(",")
        color = dsl_resolve_color(parts[0].strip())
        weight = 1.0
        if len(parts) > 1:
            try:
                weight = float(parts[1].strip())
            except ValueError:
                pass
        if color:
            shape["outline"] = {"color": color, "weight": weight}

    if "shadow" in fields:
        val = fields["shadow"].strip().lower()
        if val == "true":
            shape["shadow"] = {"offset_x": 3, "offset_y": 3, "blur": 4, "color": "#333333"}
        else:
            parts = val.split(",")
            try:
                shape["shadow"] = {
                    "offset_x": float(parts[0]) if len(parts) > 0 else 3,
                    "offset_y": float(parts[1]) if len(parts) > 1 else 3,
                    "blur": float(parts[2]) if len(parts) > 2 else 4,
                    "color": dsl_resolve_color(parts[3].strip()) if len(parts) > 3 else "#333333",
                }
            except Exception:
                pass

    for key, shape_k in (("rotation", "rotation"), ("border_radius", "border_radius"), ("line_height", "line_height")):
        if key in fields:
            try:
                shape[shape_k] = float(fields[key])
            except ValueError:
                pass

    if "valign" in fields:
        v = fields["valign"].strip().lower()
        if v in ("top", "middle", "bottom"):
            shape["vertical_align"] = v

    if "halign" in fields:
        h = fields["halign"].strip().lower()
        if h in ("left", "center", "right", "justify"):
            shape["text_align"] = h

    if "align" in fields:
        a = fields["align"].strip().lower()
        if a in ("center_x", "center_y", "center_xy"):
            shape["align"] = a

    if "font" in fields:
        shape["font"] = fields["font"].strip()

    if "padding" in fields:
        parts = fields["padding"].split(",")
        try:
            vals = [float(p.strip()) for p in parts]
            shape["padding_left"] = vals[0] if len(vals) > 0 else 0
            shape["padding_right"] = vals[1] if len(vals) > 1 else 0
            shape["padding_top"] = vals[2] if len(vals) > 2 else 0
            shape["padding_bottom"] = vals[3] if len(vals) > 3 else 0
        except Exception:
            pass

    if "bullet" in fields:
        val = fields["bullet"].strip()
        shape["bullet"] = True if val.lower() == "true" else val

    if "id" in fields:
        shape["_id"] = fields["id"]

    if text_part:
        rich = parse_rich_text(text_part)
        if rich:
            shape["rich_text"] = rich

    return shape


def build_line_shape(fields: Dict[str, str], source_line: Optional[int] = None) -> Dict[str, Any]:
    mapping = {"x1": "left", "y1": "top", "x2": "x2", "y2": "y2"}
    shape = {"type": "line", "_source_line": source_line}
    for src, dst in mapping.items():
        if src in fields:
            try:
                shape[dst] = float(fields[src])
            except ValueError:
                pass

    if "color" in fields:
        resolved = dsl_resolve_color(fields["color"])
        if resolved:
            shape["color"] = resolved

    if "weight" in fields:
        try:
            shape["line_weight"] = float(fields["weight"])
        except ValueError:
            pass

    if "dash" in fields:
        d = fields["dash"].strip().lower()
        if d in DASH_STYLES_DSL:
            shape["dash_style"] = DASH_STYLES_DSL[d]

    return shape


def build_icon_shape(fields: Dict[str, str], source_line: Optional[int] = None) -> Optional[Dict[str, Any]]:
    name = fields.get("name", "").strip()
    if not name:
        return None

    shape = {"type": "icon", "_source_line": source_line, "icon_name": name}
    style = fields.get("style", "solid").strip().lower()
    shape["icon_style"] = style if style in VALID_ICON_STYLES else "solid"

    for f in ("left", "top", "width", "height"):
        if f in fields:
            try:
                shape[f] = float(fields[f])
            except ValueError:
                pass

    if "color" in fields:
        resolved = dsl_resolve_color(fields["color"])
        if resolved:
            shape["icon_color"] = resolved

    return shape


def build_svg_shape(fields: Dict[str, str], svg_markup: str, source_line: Optional[int] = None) -> Dict[str, Any]:
    shape = {"type": "svg", "_source_line": source_line, "svg_markup": svg_markup}

    for f in ("left", "top", "width", "height"):
        if f in fields:
            try:
                shape[f] = float(fields[f])
            except ValueError:
                pass

    if "color" in fields:
        resolved = dsl_resolve_color(fields["color"])
        if resolved:
            shape["svg_color"] = resolved

    if "fit" in fields:
        fit = fields["fit"].strip().lower()
        if fit in ("contain", "cover", "stretch"):
            shape["fit"] = fit

    if "rotation" in fields:
        try:
            shape["rotation"] = float(fields["rotation"])
        except ValueError:
            pass

    if "transparency" in fields:
        try:
            shape["transparency"] = float(fields["transparency"])
        except ValueError:
            pass

    return shape


def build_table_shape(fields: Dict[str, str], subsequent_lines: List[str], source_line: Optional[int] = None) -> Dict[str, Any]:
    shape = {"type": "table", "_source_line": source_line}

    for f in ("left", "top", "width", "height"):
        if f in fields:
            try:
                shape[f] = float(fields[f])
            except ValueError:
                pass

    table_style = {}
    for key in ("header_fill", "header_text_color", "row_fill", "alt_row_fill", "text_color", "border_color"):
        if key in fields:
            resolved = dsl_resolve_color(fields[key])
            if resolved:
                table_style[key] = resolved

    if "border_weight" in fields:
        try:
            table_style["border_weight"] = float(fields["border_weight"])
        except ValueError:
            pass

    if "font" in fields:
        table_style["font"] = fields["font"].strip()

    if "font_size" in fields:
        try:
            table_style["font_size"] = float(fields["font_size"])
        except ValueError:
            pass

    if "header_bold" in fields:
        table_style["header_bold"] = fields["header_bold"].strip().lower() == "true"

    if table_style:
        shape["table_style"] = table_style

    table = {"content": [], "header_row": False}
    for line in subsequent_lines:
        line = line.strip()
        if line.startswith("cols="):
            table["col_widths"] = [float(w) for w in re.findall(r"[\d.]+", line[5:])]
        elif line.startswith("header="):
            cells = re.findall(r'"([^"]*)"', line)
            if cells:
                table["content"].insert(0, cells)
                table["header_row"] = True
        elif line.startswith("row="):
            cells = re.findall(r'"([^"]*)"', line)
            if cells:
                table["content"].append(cells)

    table["rows"] = len(table["content"])
    table["cols"] = max((len(r) for r in table["content"]), default=0)
    shape["table"] = table
    return shape


def apply_z_order_field(shape: Dict[str, Any], fields: Dict[str, str], source_line: Optional[int] = None) -> Dict[str, Any]:
    if "z_order" in fields:
        try:
            shape["z_order"] = float(fields["z_order"])
        except ValueError:
            pass
    return shape


def sort_shapes_for_render(shapes: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
    def render_key(shape):
        source_line = shape.get("_source_line", 0) or 0
        z_order = shape.get("z_order")
        if z_order is None:
            z_order = source_line * 10
        return (z_order, source_line)

    return sorted(shapes, key=render_key)


def parse_dsl(dsl_string: str) -> List[Dict[str, Any]]:
    raw_lines = dsl_string.strip().split("\n")
    shapes = []
    i = 0

    while i < len(raw_lines):
        raw = raw_lines[i]
        line = raw.strip()
        source_line = i + 1

        if not line or line.startswith("//") or line.startswith("#"):
            i += 1
            continue

        tokens, text_part = tokenize_dsl_line(line)
        if not tokens:
            i += 1
            continue

        shape_type, fields = extract_fields(tokens)
        if shape_type not in VALID_SHAPE_TYPES:
            i += 1
            continue

        if shape_type == "line":
            shape = build_line_shape(fields, source_line)
            shapes.append(apply_z_order_field(shape, fields, source_line))
            i += 1
        elif shape_type == "svg":
            svg_lines = []
            i += 1
            while i < len(raw_lines):
                nl = raw_lines[i]
                if nl.strip().lower() == "endsvg":
                    i += 1
                    break
                svg_lines.append(nl)
                i += 1
            svg_markup = "\n".join(svg_lines).strip()
            if svg_markup:
                shape = build_svg_shape(fields, svg_markup, source_line)
                shapes.append(apply_z_order_field(shape, fields, source_line))
        elif shape_type == "icon":
            parsed = build_icon_shape(fields, source_line)
            if parsed:
                shapes.append(apply_z_order_field(parsed, fields, source_line))
            i += 1
        elif shape_type == "table":
            subsequent = []
            i += 1
            while i < len(raw_lines):
                nl = raw_lines[i]
                stripped = nl.strip()
                if not stripped or stripped.split()[0] in VALID_SHAPE_TYPES:
                    break
                subsequent.append(stripped)
                i += 1
            shape = build_table_shape(fields, subsequent, source_line)
            if shape:
                shapes.append(apply_z_order_field(shape, fields, source_line))
        elif shape_type == "image":
            image_url = fields.get("url", "").strip()
            if image_url:
                shape = {
                    "type": "image",
                    "url": image_url,
                    "left": float(fields.get("left", 0)),
                    "top": float(fields.get("top", 0)),
                    "width": float(fields.get("width", 100)),
                    "height": float(fields.get("height", 100)),
                    "_source_line": source_line,
                }
                shapes.append(shape)
            i += 1
        else:
            shape = build_shape_from_fields(shape_type, fields, text_part, source_line)
            shapes.append(apply_z_order_field(shape, fields, source_line))
            i += 1

    return shapes


SLIDE_SEPARATOR = re.compile(r"^\s*---\s*$")


def parse_dsl_slides(dsl_string: str) -> List[List[Dict[str, Any]]]:
    raw_lines = dsl_string.strip().split("\n")
    blocks = []
    current = []

    for line in raw_lines:
        if SLIDE_SEPARATOR.match(line):
            blocks.append("\n".join(current))
            current = []
        else:
            current.append(line)

    if current:
        blocks.append("\n".join(current))

    result = []
    for b in blocks:
        shapes = parse_dsl(b)
        if shapes:
            result.append(shapes)

    return result if result else [[]]
