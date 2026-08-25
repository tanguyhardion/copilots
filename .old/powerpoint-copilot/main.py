import os
import sys
import hashlib
import re
import threading
import tempfile
import tkinter as tk
from tkinter import scrolledtext, messagebox
import tempfile

import requests
import urllib3

urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

# ── ICON CACHE ────────────────────────────────────────────────────────────────
ICON_CACHE_DIR = "icons"
FA_BASE_URL = (
    "https://raw.githubusercontent.com/FortAwesome/Font-Awesome/"
    "6.x/svgs/{style}/{icon}.svg"
)
VALID_ICON_STYLES = ["solid", "regular", "brands"]


def download_icon(icon_name, style="solid"):
    if style not in VALID_ICON_STYLES:
        print(f"[icon] Invalid style '{style}', falling back to 'solid'")
        style = "solid"

    style_dir = os.path.join(ICON_CACHE_DIR, style)
    os.makedirs(style_dir, exist_ok=True)
    file_path = os.path.join(style_dir, f"{icon_name}.svg")

    if os.path.exists(file_path):
        print(f"[icon] Cache hit: {file_path}")
        return file_path

    url = FA_BASE_URL.format(style=style, icon=icon_name)
    try:
        response = requests.get(url, verify=False, timeout=10)
        if response.status_code == 200:
            with open(file_path, "wb") as f:
                f.write(response.content)
            print(f"[icon] Downloaded: {file_path}")
            return file_path
        else:
            print(
                f"[icon] Not found: {icon_name} ({style}) — HTTP {response.status_code}"
            )
            for alt_style in VALID_ICON_STYLES:
                if alt_style == style:
                    continue
                alt_url = FA_BASE_URL.format(style=alt_style, icon=icon_name)
                r2 = requests.get(alt_url, verify=False, timeout=10)
                if r2.status_code == 200:
                    print(f"[icon]   '{icon_name}' available in style: '{alt_style}'")
            return None
    except Exception as e:
        print(f"[icon] Download error for '{icon_name}': {e}")
        return None


def colorize_svg(svg_path, hex_color):
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


def get_svg_aspect_ratio(svg_path):
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


def svg_to_png(svg_path, width_px=64, height_px=64):
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
    except ImportError:
        return None
    except Exception as e:
        print(f"[icon] svg→png warning: {e}")
        return None


def inject_svg_color(svg_markup, hex_color):
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


def inline_svg_to_tempfile(svg_markup, source_line=None):
    """Write inline SVG markup to a stable temp file path and return it."""
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


# ── SHAPE REGISTRY ────────────────────────────────────────────────────────────
# Decision 2: full names only — no abbreviations
VALID_SHAPE_TYPES = {
    "rect",
    "rounded_rect",
    "oval",
    "triangle",
    "star",
    "hexagon",
    "cloud",
    "arrow",
    "line",
    "text",
    "table",
    "diamond",
    "parallelogram",
    "trapezoid",
    "chevron",
    "pentagon",
    "octagon",
    "donut",
    "cross",
    "left_arrow",
    "up_arrow",
    "down_arrow",
    "left_right_arrow",
    "lightning",
    "heart",
    "frame",
    "notched_right_arrow",
    "ribbon_banner",
    "callout",
    "icon",
    "svg",
    "image",
}

# ── COLOR TABLES ──────────────────────────────────────────────────────────────
DEFAULT_DSL_COLORS = {
    "a1": "#A4D65E",
    "a2": "#6BAF5B",
    "a3": "#4A7C59",
    "a4": "#2D4A2D",
    "a5": "#7FFF00",
    "a6": "#F0F5E6",
    "bg1": "#FFFFFF",
    "bg2": "#F2F2F2",
    "t1": "#1A1A1A",
    "t2": "#4A4A4A",
}

DSL_COLOR_ALIASES = {
    "a1": "theme_accent1",
    "a2": "theme_accent2",
    "a3": "theme_accent3",
    "a4": "theme_accent4",
    "a5": "theme_accent5",
    "a6": "theme_accent6",
    "bg1": "theme_bg1",
    "bg2": "theme_bg2",
    "t1": "theme_text1",
    "t2": "theme_text2",
}

DASH_STYLES_DSL = {
    "solid": "solid",
    "dash": "dash",
    "dot": "dot",
    "dash_dot": "dash_dot",
}

SLIDE_WIDTH = 960
SLIDE_HEIGHT = 540

# ── PPT THEME ─────────────────────────────────────────────────────────────────
THEME_COLORS = {
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

PPT_THEME_MAP = {
    "theme_accent1": 5,
    "theme_accent2": 6,
    "theme_accent3": 7,
    "theme_accent4": 8,
    "theme_accent5": 9,
    "theme_accent6": 10,
    "theme_text1": 13,
    "theme_text2": 14,
    "theme_bg1": 12,
    "theme_bg2": 15,
}

MSO_SHAPE_MAP = {
    "rect": 1,
    "oval": 9,
    "triangle": 7,
    "rounded_rect": 5,
    "star": 92,
    "hexagon": 10,
    "cloud": 179,
    "arrow": 33,
    "diamond": 4,
    "parallelogram": 2,
    "trapezoid": 3,
    "chevron": 52,
    "pentagon": 51,
    "octagon": 6,
    "donut": 18,
    "cross": 11,
    "left_arrow": 34,
    "up_arrow": 35,
    "down_arrow": 36,
    "left_right_arrow": 37,
    "lightning": 22,
    "heart": 21,
    "frame": 158,
    "notched_right_arrow": 50,
    "ribbon_banner": 53,
    "callout": 106,
}


# ── COLOR HELPERS ─────────────────────────────────────────────────────────────
def dsl_resolve_color(c):
    """
    Resolve DSL color strings into either:
      - a PowerPoint theme token (for base aliases like a1, t1, bg1)
      - a derived hex value (for variants like a1_l1, a1_d2)
      - a raw hex color
    """
    if not c:
        return None

    c = c.strip().lower()

    # Base theme aliases remain theme-bound
    if c in DSL_COLOR_ALIASES:
        return DSL_COLOR_ALIASES[c]

    # Theme variants: a1_l1 / a1_l2 / a1_d1 / a1_d2
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

    # Short hex
    if re.match(r"^#[0-9a-f]{3}$", c):
        return "#" + "".join(ch * 2 for ch in c[1:])

    # Full hex
    if re.match(r"^#[0-9a-f]{6}$", c):
        return c.upper()

    return None


def hex_to_rgb(hex_color):
    try:
        hex_color = re.sub(r"[^0-9A-Fa-f]", "", hex_color.strip().lstrip("#"))
        if len(hex_color) == 3:
            hex_color = "".join(c * 2 for c in hex_color)
        if len(hex_color) != 6:
            raise ValueError(f"bad hex length: '{hex_color}'")
        return tuple(int(hex_color[i : i + 2], 16) for i in (0, 2, 4))
    except Exception as e:
        print(f"hex_to_rgb warning: {e}")
        return (128, 128, 128)


def rgb_to_bgr_int(r, g, b):
    return r + (g << 8) + (b << 16)


def is_theme_color(c):
    return c in PPT_THEME_MAP


def clamp(v, lo=0, hi=255):
    return max(lo, min(hi, int(round(v))))


def rgb_to_hex(r, g, b):
    return f"#{clamp(r):02X}{clamp(g):02X}{clamp(b):02X}"


def lighten_hex(hex_color, amount):
    """
    Mix a color toward white by amount in [0,1].
    """
    r, g, b = hex_to_rgb(hex_color)
    nr = r + (255 - r) * amount
    ng = g + (255 - g) * amount
    nb = b + (255 - b) * amount
    return rgb_to_hex(nr, ng, nb)


def darken_hex(hex_color, amount):
    """
    Mix a color toward black by amount in [0,1].
    """
    r, g, b = hex_to_rgb(hex_color)
    nr = r * (1 - amount)
    ng = g * (1 - amount)
    nb = b * (1 - amount)
    return rgb_to_hex(nr, ng, nb)


def bgr_int_to_hex(rgb_int):
    """
    PowerPoint COM returns colors as Office RGB integers, which are actually
    stored BGR-style for easy assignment via .RGB.
    """
    try:
        v = int(rgb_int)
        r = v & 0xFF
        g = (v >> 8) & 0xFF
        b = (v >> 16) & 0xFF
        return f"#{r:02X}{g:02X}{b:02X}"
    except Exception as e:
        print(f"bgr_int_to_hex warning: {e}")
        return None


def get_active_ppt_theme_colors():
    """
    Read the active presentation's base theme colors from PowerPoint and return
    a dict keyed by theme token name, e.g. 'theme_accent1' -> '#123456'.

    Falls back to the default theme-color hexes if anything fails.
    """
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
            if ppt.Presentations.Count == 0:
                return colors

            pres = ppt.ActivePresentation
            scheme = pres.SlideMaster.Theme.ThemeColorScheme

            for theme_name, idx in PPT_THEME_MAP.items():
                try:
                    office_color = scheme.Colors(idx)
                    hex_value = bgr_int_to_hex(office_color.RGB)
                    if hex_value:
                        colors[theme_name] = hex_value
                except Exception as inner_e:
                    print(
                        f"[theme] failed reading {theme_name} at index {idx}: {inner_e}"
                    )

        finally:
            pythoncom.CoUninitialize()

    except Exception as e:
        print(f"[theme] using fallback defaults: {e}")

    return colors


def refresh_dsl_theme_colors():
    """
    Refresh THEME_COLORS from the active PowerPoint presentation.
    Base aliases stay semantic (a1 -> theme_accent1, etc.).
    """
    global THEME_COLORS

    THEME_COLORS = get_active_ppt_theme_colors()

    print("[theme] Theme colors refreshed from active PowerPoint theme:")
    for k, v in THEME_COLORS.items():
        print(f"  {k} = {v}")


def resolve_alignment(shape_def):
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


# ── DSL TOKENIZER ─────────────────────────────────────────────────────────────
def tokenize_dsl_line(line):
    """
    Split a single DSL line into:
      - before_pipe tokens  (list of 'key=value' or bare words)
      - after_pipe string   (raw text part, or None)

    Handles quoted values that may contain spaces or = signs.
    """
    # strip inline comments
    # Strip inline comments: // must be preceded by whitespace or start-of-line
    # This prevents stripping https:// or similar URL protocols
    line = re.sub(r"(?<!\S)//.*$", "", line).strip()
    if not line:
        return [], None

    # split on the FIRST pipe that is not inside quotes
    before, after = _split_on_pipe(line)

    tokens = _tokenize_kvs(before)
    return tokens, after.strip() if after is not None else None


def _split_on_pipe(s):
    """Return (before, after) split on first unquoted |. after is None if no |."""
    in_quotes = False
    for i, ch in enumerate(s):
        if ch == '"':
            in_quotes = not in_quotes
        elif ch == "|" and not in_quotes:
            return s[:i], s[i + 1 :]
    return s, None


def _tokenize_kvs(s):
    """
    Tokenize 'key=value' pairs where value may be quoted.
    Returns a list of raw token strings like ['rect', 'left=20', 'color=a1'].
    """
    tokens = []
    pattern = re.compile(
        r"(\w[\w\-]*"  # key (word chars + hyphens for icon names)
        r'(?:=[^\s"]*|'  # =non-space-value
        r'="[^"]*")?'  # or ="quoted value"
        r")"
    )
    for m in pattern.finditer(s):
        tok = m.group(1).strip()
        if tok:
            tokens.append(tok)
    return tokens


def extract_fields(tokens):
    """
    Given token list like ['rect', 'left=20', 'top=70'],
    returns (shape_type: str, fields: dict).
    """
    if not tokens:
        return None, {}
    shape_type = tokens[0].lower()
    fields = {}
    for tok in tokens[1:]:
        if "=" in tok:
            key, _, value = tok.partition("=")
            # strip surrounding quotes from value if present
            value = value.strip().strip('"')
            fields[key.strip().lower()] = value
    return shape_type, fields


# ── TEXT SEGMENT PARSER ───────────────────────────────────────────────────────
def parse_text_segment(seg):
    """
    Parses one rich-text segment.

    New format (Decision 7):
        "text content" size=12 bold=true italic=true underline=true color=#FFF

    Returns a dict or None.
    """
    seg = seg.strip()
    m = re.match(r'"((?:[^"\\]|\\.)*)"(.*)', seg)
    if not m:
        return None

    text = m.group(1).replace("\\n", "\n").replace('\\"', '"')
    rest = m.group(2).strip()
    result = {"text": text}

    if not rest:
        return result

    # ── new key=value format ──────────────────────────────────────────────────
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


def parse_rich_text(text_part):
    """Split on + (outside quotes) and parse each segment."""
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


# ── FIELD NORMALIZER ──────────────────────────────────────────────────────────
def build_shape_from_fields(shape_type, fields, text_part, source_line=None):
    """
    Convert raw string field-dict into a normalized shape dict.
    Every supported field is explicitly handled here.
    """
    shape = {"type": shape_type}
    if source_line is not None:
        shape["_source_line"] = source_line

    # ── geometry ──────────────────────────────────────────────────────────────
    for f in ("left", "top", "width", "height"):
        if f in fields:
            try:
                shape[f] = float(fields[f])
            except ValueError:
                print(f"[dsl] line {source_line}: bad number for '{f}': {fields[f]}")

    # ── fill / transparency ───────────────────────────────────────────────────
    if "color" in fields:
        resolved = dsl_resolve_color(fields["color"])
        if resolved:
            shape["color"] = resolved
        else:
            print(f"[dsl] line {source_line}: unresolved color '{fields['color']}'")

    if "transparency" in fields:
        try:
            shape["transparency"] = float(fields["transparency"])
        except ValueError:
            pass

    # ── gradient (Decision 4.B) ───────────────────────────────────────────────
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

    # ── outline (Decision 6) ──────────────────────────────────────────────────
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

    # ── shadow (Decision 5.B) ─────────────────────────────────────────────────
    if "shadow" in fields:
        val = fields["shadow"].strip().lower()
        if val == "true":
            shape["shadow"] = {
                "offset_x": 3,
                "offset_y": 3,
                "blur": 4,
                "color": "#333333",
            }
        else:
            parts = val.split(",")
            try:
                shape["shadow"] = {
                    "offset_x": float(parts[0]) if len(parts) > 0 else 3,
                    "offset_y": float(parts[1]) if len(parts) > 1 else 3,
                    "blur": float(parts[2]) if len(parts) > 2 else 4,
                    "color": (
                        dsl_resolve_color(parts[3].strip())
                        if len(parts) > 3
                        else "#333333"
                    ),
                }
            except ValueError as e:
                print(f"[dsl] line {source_line}: bad shadow value — {e}")

    # ── rotation / border_radius ──────────────────────────────────────────────
    if "rotation" in fields:
        try:
            shape["rotation"] = float(fields["rotation"])
        except ValueError:
            pass

    if "border_radius" in fields:
        try:
            shape["border_radius"] = float(fields["border_radius"])
        except ValueError:
            pass

    # ── text layout (Decision 8) ──────────────────────────────────────────────
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

    if "line_height" in fields:
        try:
            shape["line_height"] = float(fields["line_height"])
        except ValueError:
            pass

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
        except ValueError as e:
            print(f"[dsl] line {source_line}: bad padding — {e}")

    if "bullet" in fields:
        val = fields["bullet"].strip()
        shape["bullet"] = True if val.lower() == "true" else val

    # ── optional debug id ─────────────────────────────────────────────────────
    if "id" in fields:
        shape["_id"] = fields["id"]

    # ── text content (Decision 7) ─────────────────────────────────────────────
    if text_part:
        rich = parse_rich_text(text_part)
        if rich:
            shape["rich_text"] = rich

    return shape


# ── SHAPE-SPECIFIC BUILDERS ───────────────────────────────────────────────────
def build_line_shape(fields, source_line=None):
    """
    line x1=<n> y1=<n> x2=<n> y2=<n> [color=<c>] [weight=<n>] [dash=solid|dash|dot|dash_dot]
    """
    shape = {"type": "line"}
    if source_line is not None:
        shape["_source_line"] = source_line

    for f in ("x1", "y1", "x2", "y2"):
        if f in fields:
            try:
                shape[f if f in ("x2", "y2") else ("left" if f == "x1" else "top")] = (
                    float(fields[f])
                )
                # map x1→left, y1→top, keep x2/y2 as-is
            except ValueError:
                pass

    # fix: map explicitly so we don't confuse ourselves
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


def build_icon_shape(fields, source_line=None):
    """
    icon name=<n> style=solid|regular|brands left=<n> top=<n> width=<n> height=<n> [color=<c>]
    """
    shape = {"type": "icon", "_source_line": source_line}

    name = fields.get("name", "").strip()
    if not name:
        print(f"[dsl] line {source_line}: icon missing 'name='")
        return None
    shape["icon_name"] = name

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


def build_svg_shape(fields, svg_markup, source_line=None):
    """
    svg left=<n> top=<n> width=<n> height=<n> [color=<c>] [fit=contain|cover|stretch] [rotation=<n>]
      <svg ...>
      ...
      </svg>
      endsvg
    """
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


def build_table_shape(fields, subsequent_lines, source_line=None):
    """
    table left=<n> top=<n> width=<n> height=<n>
            [header_fill=<c>] [header_text_color=<c>] [header_bold=true|false]
            [row_fill=<c>] [alt_row_fill=<c>] [text_color=<c>]
            [border_color=<c>] [border_weight=<n>] [font=Name] [font_size=<n>]
      cols=<w1>,<w2>,...
      header="col1","col2",...
      row="val1","val2",...
    """
    shape = {"type": "table", "_source_line": source_line}

    for f in ("left", "top", "width", "height"):
        if f in fields:
            try:
                shape[f] = float(fields[f])
            except ValueError:
                pass

    table_style = {}
    for key in (
        "header_fill",
        "header_text_color",
        "row_fill",
        "alt_row_fill",
        "text_color",
        "border_color",
    ):
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

    if "header_text_bold" in fields:
        table_style["header_bold"] = (
            fields["header_text_bold"].strip().lower() == "true"
        )

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


def apply_z_order_field(shape, fields, source_line=None):
    """Attach an explicit z-order value when the DSL provides one."""
    if "z_order" not in fields:
        return shape

    try:
        shape["z_order"] = float(fields["z_order"])
    except ValueError:
        print(f"[dsl] line {source_line}: bad z_order value '{fields['z_order']}'")
    return shape


def sort_shapes_for_render(shapes):
    """
    Sort shapes from back to front.

    Shapes without an explicit z_order keep their original DSL order by using
    their source line as the default layer rank.
    """

    def render_key(shape):
        source_line = shape.get("_source_line", 0) or 0
        z_order = shape.get("z_order")
        if z_order is None:
            z_order = source_line * 10
        return (z_order, source_line)

    return sorted(shapes, key=render_key)


# ── MAIN PARSER ───────────────────────────────────────────────────────────────
def parse_dsl(dsl_string):
    raw_lines = dsl_string.strip().split("\n")
    shapes = []
    i = 0

    while i < len(raw_lines):
        raw = raw_lines[i]
        line = raw.strip()
        source_line = i + 1

        # blank / comment
        if not line or line.startswith("//") or line.startswith("#"):
            i += 1
            continue

        tokens, text_part = tokenize_dsl_line(line)
        if not tokens:
            i += 1
            continue

        shape_type, fields = extract_fields(tokens)

        # ── validate shape type ───────────────────────────────────────────────
        if shape_type not in VALID_SHAPE_TYPES:
            print(f"[dsl] line {source_line}: unknown shape '{shape_type}' — skipped")
            i += 1
            continue

        # ── dispatch ──────────────────────────────────────────────────────────
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
            if not svg_markup:
                print(f"[dsl] line {source_line}: empty svg block — skipped")
                continue
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

                # Stop if blank or next shape definition
                if not stripped or stripped.split()[0] in VALID_SHAPE_TYPES:
                    break

                subsequent.append(stripped)
                i += 1

            shape = build_table_shape(fields, subsequent, source_line)
            if shape:
                shapes.append(apply_z_order_field(shape, fields, source_line))

        elif shape_type == "image":
            image_url = fields.get("url", "").strip()
            if not image_url:
                print(f"[dsl] line {source_line}: image missing 'url='")
                i += 1
                continue
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


# ── MULTI-SLIDE PARSER ────────────────────────────────────────────────────────
SLIDE_SEPARATOR = re.compile(r"^\s*---\s*$")


def parse_dsl_slides(dsl_string):
    """
    Split *dsl_string* on slide-separator lines (lines that contain only '---')
    and parse each block independently.

    Returns a list of slide-shape-lists:
        [ [shape, shape, …],   # slide 1
          [shape, shape, …],   # slide 2
          … ]

    A single block with no separator is returned as a one-element list,
    so all existing callers that just want shapes[0] still work.
    """
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

    # Drop completely empty blocks (e.g. two consecutive ---)
    result = []
    for block in blocks:
        if block.strip():
            result.append(parse_dsl(block))

    return result if result else [[]]


# ── SAMPLE DSL ────────────────────────────────────────────────────────────────
SAMPLE_DSL = """\
// Header
rect left=0 top=0 width=960 height=60 color=a4 | "Simple DSL Demo" size=22 bold=true color=#FFFFFF

// Basic shapes
rect left=40 top=90 width=180 height=80 color=a1 | "Rectangle" size=12 bold=true color=t1
rounded_rect left=250 top=90 width=180 height=80 color=a2 outline=#FFFFFF,2 | "Rounded" size=12 bold=true color=#FFFFFF
oval left=460 top=90 width=140 height=80 color=a3 | "Oval" size=12 bold=true color=#FFFFFF
triangle left=640 top=90 width=120 height=80 color=a2 | "Triangle" size=12 bold=true color=#FFFFFF

// Gradient
rect left=40 top=210 width=220 height=90 color=a1 grad_stops=a1,a4 grad_angle=90 | "Gradient" size=12 bold=true color=#FFFFFF

// Text alignment
rect left=290 top=210 width=220 height=90 color=a6 outline=a3,1 valign=middle halign=center | "Centered Text" size=12 bold=true color=a3

// Rich text
rect left=540 top=210 width=300 height=90 color=a6 outline=a3,1 padding=12,12,8,8 | "Bold " size=13 bold=true color=a4 + "Normal " size=13 color=t1 + "Italic" size=13 italic=true color=a3

// Line styles
line x1=40 y1=340 x2=300 y2=340 color=a3 weight=2 dash=solid
line x1=40 y1=360 x2=300 y2=360 color=a2 weight=2 dash=dash
line x1=40 y1=380 x2=300 y2=380 color=a1 weight=2 dash=dot

// Icon
icon name=shield-halved style=solid left=360 top=330 width=50 height=50 color=a1

// Table
table left=460 top=330 width=360 height=140
cols=120,100,140
header="Feature","Status","Notes"
row="Shapes","Done","Basic set"
row="Text","Done","Rich text"
row="Icons","Done","Font Awesome"

// Footer line
line x1=0 y1=530 x2=960 y2=530 color=a4 weight=1 dash=solid
"""


# ── UTF-16 length helper ──────────────────────────────────────────────────────
def utf16_len(text: str) -> int:
    """
    Return the number of UTF-16 code units in *text*.
    Characters outside the BMP (e.g. most emoji, U+10000+) each
    occupy TWO code units, which is how PowerPoint / COM counts them.
    """
    return len(text.encode("utf-16-le")) // 2


# ── POWERPOINT CONNECTOR ──────────────────────────────────────────────────────
class PowerPointConnector:
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

    def _prefetch_icons(self, shapes, status_cb=None):
        icon_shapes = [s for s in shapes if s.get("type") == "icon"]
        if not icon_shapes:
            return {}
        cache = {}
        total = len(icon_shapes)
        for i, s in enumerate(icon_shapes):
            name = s["icon_name"]
            style = s.get("icon_style", "solid")
            key = (name, style)
            if key not in cache:
                if status_cb:
                    status_cb(f"Downloading icon {i+1}/{total}: {name} ({style})…")
                cache[key] = download_icon(name, style)
        return cache

    # ── copy to clipboard ─────────────────────────────────────────────────────
    def create_shapes_and_copy(self, shapes, status_cb=None):
        import win32com.client
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

    def create_on_new_slide(self, shapes_or_slides, status_cb=None):
        """
        Create one new slide per slide-block.

        Accepts:
        - list[dict]        → single slide (backward-compatible)
        - list[list[dict]]  → one slide per inner list
        """
        import pythoncom

        pythoncom.CoInitialize()
        try:
            # ── normalise input ───────────────────────────────────────────────
            if shapes_or_slides and isinstance(shapes_or_slides[0], dict):
                # Old-style single list → wrap in a list-of-one
                slides_data = [shapes_or_slides]
            else:
                slides_data = shapes_or_slides  # already list-of-lists

            # ── pre-fetch all icons across all slides in one pass ─────────────
            all_shapes = [s for slide in slides_data for s in slide]
            icon_cache = self._prefetch_icons(all_shapes, status_cb)

            self.connect()
            ppt = self.ppt_app
            if ppt.Presentations.Count == 0:
                raise Exception(
                    "No presentation open. Please open one in PowerPoint first."
                )
            pres = ppt.ActivePresentation

            try:
                current_index = ppt.ActiveWindow.View.Slide.SlideIndex
            except Exception:
                current_index = pres.Slides.Count

            total_slides = len(slides_data)

            for slide_num, shapes in enumerate(slides_data):
                new_index = current_index + slide_num + 1
                if status_cb:
                    status_cb(
                        f"Creating slide {slide_num + 1}/{total_slides} "
                        f"(index {new_index})…"
                    )
                slide = pres.Slides.Add(new_index, 12)

                # ── Switch to this slide BEFORE adding shapes ────────────────────
                try:
                    ppt.ActiveWindow.View.GotoSlide(new_index)
                    ppt.ActiveWindow.Activate()  # Bring PowerPoint window to front
                except Exception as e:
                    print(f"[ppt] Could not navigate to slide {new_index}: {e}")

                ordered = sort_shapes_for_render(shapes)
                total_shapes = len(ordered)
                for i, sd in enumerate(ordered):
                    if status_cb:
                        status_cb(
                            f"Slide {slide_num + 1}/{total_slides} — "
                            f"shape {i + 1}/{total_shapes}…"
                        )
                    self._create_single_shape(slide, sd, icon_cache)

            if status_cb:
                status_cb(
                    f"✓ {total_slides} slide(s) created "
                    f"({sum(len(s) for s in slides_data)} shapes total)!"
                )
        finally:
            pythoncom.CoUninitialize()

    # ── single shape dispatcher ───────────────────────────────────────────────
    def _create_single_shape(self, slide, shape_def, icon_cache=None):
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

        ppt_shape = None
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
                        ppt_shape.Line.ForeColor.ObjectThemeColor = PPT_THEME_MAP[
                            ol_color
                        ]
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
                            try:
                                ppt_shape.Adjustments._oleobj_.Invoke(
                                    0, 0, 4, 0, 1, val
                                )
                            except Exception as e:
                                print(f"border_radius warning: {e}")

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

    # ── icon ──────────────────────────────────────────────────────────────────
    def _create_icon(self, slide, shape_def, icon_cache):
        name = shape_def.get("icon_name", "question")
        style = shape_def.get("icon_style", "solid")
        left = shape_def.get("left", 0)
        top = shape_def.get("top", 0)
        width = shape_def.get("width", 48)
        height = shape_def.get("height", 48)
        color = shape_def.get("icon_color")

        svg_path = icon_cache.get((name, style))
        if not svg_path or not os.path.exists(svg_path):
            print(f"[icon] Missing file for '{name}' ({style}) — inserting placeholder")
            return self._icon_placeholder(slide, shape_def)

        working_path = colorize_svg(svg_path, color) if color else svg_path
        aspect_ratio = get_svg_aspect_ratio(working_path)
        if aspect_ratio <= 0:
            aspect_ratio = 1.0

        final_width = width
        final_height = width / aspect_ratio

        png_path = svg_to_png(
            working_path,
            width_px=max(1, int(final_width * 2)),
            height_px=max(1, int(final_height * 2)),
        )
        insert_path = (
            png_path if png_path and os.path.exists(png_path) else working_path
        )
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
        except Exception as e:
            print(f"[icon] AddPicture failed for '{name}': {e}")
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
            print(
                f"[svg] Missing SVG markup on line {shape_def.get('_source_line', '?')}"
            )
            return None

        if svg_color:
            svg_markup = inject_svg_color(svg_markup, svg_color)

        svg_path = inline_svg_to_tempfile(svg_markup, shape_def.get("_source_line"))
        if not svg_path or not os.path.exists(svg_path):
            print(
                f"[svg] Unable to create temp file for line {shape_def.get('_source_line', '?')}"
            )
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
        except Exception as e:
            print(
                f"[svg] AddPicture failed on line {shape_def.get('_source_line', '?')}: {e}"
            )
            png_path = svg_to_png(
                svg_path,
                width_px=max(1, int(render_width * 2)),
                height_px=max(1, int(render_height * 2)),
            )
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
                except Exception as inner_e:
                    print(f"[svg] PNG fallback failed: {inner_e}")
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

    # ── shadow ────────────────────────────────────────────────────────────────
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
        except Exception as e:
            print(f"Shadow warning: {e}")

    # ── gradient ──────────────────────────────────────────────────────────────
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

        except Exception as e:
            print(f"Gradient warning: {e}")
            try:
                r, g, b = hex_to_rgb(stops[0])
                ppt_shape.Fill.ForeColor.RGB = rgb_to_bgr_int(r, g, b)
                ppt_shape.Fill.Solid()
            except Exception:
                pass

    # ── table ─────────────────────────────────────────────────────────────────
    def _create_table(self, slide, shape_def):
        left = shape_def.get("left", 0)
        top = shape_def.get("top", 0)
        width = shape_def.get("width", 400)
        height = shape_def.get("height", 200)
        td = shape_def.get("table", {})
        style = shape_def.get("table_style", {})
        rows = td.get("rows", 2)  # tofix
        cols = td.get("cols", 2)  # tofix

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
                            cell_shape.Fill.ForeColor.ObjectThemeColor = PPT_THEME_MAP[
                                fill_color
                            ]
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
                                cell_shape.Line.ForeColor.ObjectThemeColor = (
                                    PPT_THEME_MAP[border_color]
                                )
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

                        text_color = (
                            style.get("header_text_color")
                            if is_header and style.get("header_text_color")
                            else style.get("text_color")
                        )
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

    # ── line ──────────────────────────────────────────────────────────────────
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
        dash_map = {
            "solid": 1,
            "dash": 4,
            "dot": 3,
            "dash_dot": 5,
        }
        ppt_shape.Line.DashStyle = dash_map.get(shape_def.get("dash_style", "solid"), 1)
        return ppt_shape

    def _create_image(self, slide, shape_def):
        url = shape_def.get("url")
        try:
            headers = {
                "User-Agent": (
                    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
                    "(KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36"
                )
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

            # Insert at natural size first
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

                # Unlock first so both dimensions can be set independently
                pic.LockAspectRatio = False
                pic.Width = final_w
                pic.Height = final_h
                pic.LockAspectRatio = True

                # Center inside the requested bounding box
                pic.Left = desired_left + (desired_width - final_w) / 2
                pic.Top = desired_top + (desired_height - final_h) / 2

            if rotation:
                pic.Rotation = rotation

            return pic

        except Exception as e:
            print(f"[image] Failed to download {url}: {e}")
            return None

    # ── rich text ─────────────────────────────────────────────────────────────
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
        tf.VerticalAnchor = v_map.get(
            shape_def.get("vertical_align", default_v), v_map[default_v]
        )

        # ── build runs one-by-one instead of slicing a pre-set string ────────────
        # Clear existing text first
        tf.TextRange.Text = ""

        for idx, seg in enumerate(rich_text):
            text = seg.get("text", "")
            if not text:
                continue

            if idx == 0:
                # First segment: use the range that already exists after clearing
                tf.TextRange.Text = text
                tr = tf.TextRange
            else:
                # Subsequent segments: insert after the current last character
                end_pos = utf16_len(tf.TextRange.Text)
                tr = tf.TextRange.InsertAfter(text)

            # ── apply formatting to this run ──────────────────────────────────
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

        # ── paragraph-level formatting ────────────────────────────────────────────
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
                    if lh_val <= 0:
                        raise ValueError(f"line_height must be > 0, got {lh_val}")
                    para.ParagraphFormat.SpaceWithin = lh_val
                except Exception as e:
                    print(f"Line-height warning: {e}")

    # ── simple text ───────────────────────────────────────────────────────────
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
        tf.VerticalAnchor = v_map.get(
            shape_def.get("vertical_align", default_v), v_map[default_v]
        )

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
                    if lh_val <= 0:
                        raise ValueError(f"line_height must be > 0, got {lh_val}")
                    para.ParagraphFormat.SpaceWithin = lh_val
                except Exception as e:
                    print(f"Line-height warning: {e}")

    def create_on_current_slide(self, shapes, status_cb=None):
        """Insert shapes on the currently active slide (no new slide created)"""
        import pythoncom

        pythoncom.CoInitialize()
        try:
            icon_cache = self._prefetch_icons(shapes, status_cb)
            self.connect()
            ppt = self.ppt_app

            if ppt.Presentations.Count == 0:
                raise Exception(
                    "No presentation open. Please open one in PowerPoint first."
                )

            try:
                slide = ppt.ActiveWindow.View.Slide
                slide_index = slide.SlideIndex
            except Exception:
                raise Exception(
                    "Could not get active slide. Please click on a slide first."
                )

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


# ── SYNTAX HIGHLIGHTING CONFIG ────────────────────────────────────────────────
# All valid keywords the highlighter needs to know about
HL_SHAPE_NAMES = sorted(VALID_SHAPE_TYPES, key=len, reverse=True)

HL_FIELD_KEYS = [
    "left",
    "top",
    "width",
    "height",
    "color",
    "transparency",
    "grad_stops",
    "grad_angle",
    "outline",
    "shadow",
    "rotation",
    "border_radius",
    "z_order",
    "valign",
    "halign",
    "align",
    "line_height",
    "padding",
    "bullet",
    "id",
    # line-specific
    "x1",
    "y1",
    "x2",
    "y2",
    "weight",
    "dash",
    # icon-specific
    "name",
    "style",
    "fit",
    # text segment
    "size",
    "bold",
    "italic",
    "underline",
    "font",
    # table styling
    "header_fill",
    "header_text_color",
    "header_bold",
    "header_text_bold",
    "row_fill",
    "alt_row_fill",
    "text_color",
    "border_color",
    "border_weight",
    "font_size",
]

HL_COLOR_ALIASES = list(DSL_COLOR_ALIASES.keys())

def resource_path(relative_path):
    if hasattr(sys, "_MEIPASS"):
        return os.path.join(sys._MEIPASS, relative_path)
    return os.path.join(os.path.abspath("."), relative_path)

# ── GUI ───────────────────────────────────────────────────────────────────────
class App:
    def __init__(self):
        self.root = tk.Tk()
        self.root.title("PowerPoint Copilot")
        self.root.state("zoomed")
        self.root.configure(bg="#1E1E2E")
        icon_path = resource_path("assets/icon.ico")
        self.root.iconbitmap(icon_path)
        self.connector = PowerPointConnector()
        self._build()

    def _build(self):
        # palette
        BG_MAIN = "#1E1E2E"
        BG_PANEL = "#2B2B2B"
        FG_MAIN = "#F3F3F3"
        FG_MUTED = "#B8B8B8"

        PPT_BLUE = "#4472C4"
        PPT_BLUE_HOVER = "#5B84D6"

        PPT_ORANGE = "#ED7D31"
        PPT_ORANGE_HOVER = "#F08F4F"

        PPT_GRAYBTN = "#4A4F57"
        PPT_GRAYBTN_HOVER = "#5A606A"

        # ── header ────────────────────────────────────────────────────────────
        header = tk.Frame(self.root, bg=BG_PANEL, height=48)
        header.pack(fill="x")
        header.pack_propagate(False)

        tk.Label(
            header,
            text="⬡  PowerPoint Copilot",
            font=("Segoe UI", 14, "bold"),
            bg=BG_PANEL,
            fg=FG_MAIN,
        ).pack(side="left", padx=16, pady=10)

        # ── editor ────────────────────────────────────────────────────────────
        editor_frame = tk.Frame(self.root, bg=BG_MAIN)
        editor_frame.pack(fill="both", expand=True, padx=10, pady=(8, 4))

        tk.Label(
            editor_frame,
            text="Paste DSL below:",
            font=("Segoe UI", 9),
            bg=BG_MAIN,
            fg=FG_MUTED,
        ).pack(anchor="w")

        self.editor = scrolledtext.ScrolledText(
            editor_frame,
            wrap="word",
            font=("Consolas", 10),
            bg=BG_MAIN,
            fg="#CDD6F4",
            insertbackground=PPT_BLUE,
            selectbackground="#3E5F99",
            selectforeground="#FFFFFF",
            relief="flat",
            bd=8,
            undo=True,
        )

        self.editor.pack(fill="both", expand=True)
        self.editor.insert("1.0", SAMPLE_DSL)
        self._setup_highlighting()

        # ── button row ────────────────────────────────────────────────────────
        btn_row = tk.Frame(self.root, bg=BG_MAIN)
        btn_row.pack(fill="x", padx=10, pady=(4, 6))

        self.copy_btn = tk.Button(
            btn_row,
            text="📋  Copy to Clipboard",
            font=("Segoe UI", 11, "bold"),
            bg=PPT_GRAYBTN,
            fg="white",
            activebackground=PPT_GRAYBTN_HOVER,
            activeforeground="white",
            relief="flat",
            bd=0,
            padx=18,
            pady=8,
            cursor="hand2",
            command=self.copy_to_clipboard,
        )
        self.copy_btn.pack(side="left", padx=(0, 8))

        self.insert_btn = tk.Button(
            btn_row,
            text="➕  Insert on Current Slide",
            font=("Segoe UI", 11, "bold"),
            bg=PPT_BLUE,
            fg="white",
            activebackground=PPT_BLUE_HOVER,
            activeforeground="white",
            relief="flat",
            bd=0,
            padx=18,
            pady=8,
            cursor="hand2",
            command=self.insert_on_current_slide,
        )
        self.insert_btn.pack(side="left", padx=(0, 8))

        self.slide_btn = tk.Button(
            btn_row,
            text="📄  Create Slide(s)",
            font=("Segoe UI", 11, "bold"),
            bg=PPT_ORANGE,
            fg="white",
            activebackground=PPT_ORANGE_HOVER,
            activeforeground="white",
            relief="flat",
            bd=0,
            padx=18,
            pady=8,
            cursor="hand2",
            command=self.create_full_slide,
        )
        self.slide_btn.pack(side="left", padx=(0, 8))

        # ── status bar ────────────────────────────────────────────────────────
        self.status = tk.Label(
            self.root,
            text="Ready: paste DSL and hit a button",
            font=("Segoe UI", 9),
            bg=BG_PANEL,
            fg=FG_MUTED,
            anchor="w",
        )
        self.status.pack(fill="x", side="bottom", ipady=5, padx=10)

        self.root.after(100, self._highlight)

    # ── syntax highlighting ───────────────────────────────────────────────────
    def _setup_highlighting(self):
        self.editor.tag_configure("hl_shape", foreground="#5B9BD5")  # Office blue
        self.editor.tag_configure("hl_key", foreground="#9E7CC1")  # muted violet
        self.editor.tag_configure("hl_value", foreground="#ED7D31")  # PPT orange
        self.editor.tag_configure("hl_color", foreground="#70AD47")  # Office green
        self.editor.tag_configure("hl_string", foreground="#FFD966")  # soft gold
        self.editor.tag_configure(
            "hl_comment",
            foreground="#7F7F7F",
            font=("Consolas", 10, "italic"),
        )
        self.editor.tag_configure("hl_pipe", foreground="#A5A5A5")
        self.editor.tag_configure("hl_equals", foreground="#808080")

        # SVG/XML-specific
        self.editor.tag_configure("hl_svg_tag", foreground="#5B9BD5")
        self.editor.tag_configure("hl_svg_attr", foreground="#9E7CC1")
        self.editor.tag_configure("hl_svg_comment", foreground="#7F7F7F")
        self.editor.tag_configure("hl_svg_endsvg", foreground="#ED7D31")
        self.editor.tag_configure(
            "hl_separator",
            foreground="#FFC000",
            font=("Consolas", 10, "bold"),
        )

        self.editor.bind("<KeyRelease>", self._highlight)

    def _highlight(self, event=None):
        # ── build patterns ────────────────────────────────────────────────────
        shape_pattern = (
            r"^(" + "|".join(re.escape(s) for s in HL_SHAPE_NAMES) + r")(?=\s|$)"
        )

        key_pattern = (
            r"(?<!\w)(" + "|".join(re.escape(k) for k in HL_FIELD_KEYS) + r")(?==)"
        )

        color_alias_pattern = (
            r"(?<==)("
            + "|".join(re.escape(c) for c in HL_COLOR_ALIASES)
            + r")(?=\s|$|,)"
        )

        color_variant_pattern = r"(?<==)(a[1-6]|bg[12]|t[12])_(l1|l2|d1|d2)(?=\s|$|,)"
        hex_color_pattern = r"#[0-9A-Fa-f]{3,6}\b"
        number_pattern = r"(?<==)-?\d+(\.\d+)?"
        string_pattern = r'"[^"]*"'
        pipe_pattern = r"\|"
        equals_pattern = r"="

        # SVG/XML patterns
        svg_tag_pattern = r"</?[A-Za-z_:][\w:.\-]*"
        svg_attr_pattern = r"(?<=\s)([A-Za-z_:][\w:.\-]*)(?=\=)"
        svg_comment_pattern = r"<!--.*?-->"
        endsvg_pattern = r"^\s*endsvg\s*$"

        all_tags = (
            "hl_shape",
            "hl_key",
            "hl_value",
            "hl_color",
            "hl_string",
            "hl_comment",
            "hl_pipe",
            "hl_equals",
            "hl_svg_tag",
            "hl_svg_attr",
            "hl_svg_comment",
            "hl_svg_endsvg",
            "hl_separator",
        )

        for tag in all_tags:
            self.editor.tag_remove(tag, "1.0", "end")

        content = self.editor.get("1.0", "end-1c")
        in_svg_block = False

        for li, line in enumerate(content.split("\n")):
            row = li + 1

            # Build a mapping from Python string offset to Tkinter column offset.
            # Tkinter counts surrogate pairs (emoji etc.) as 2 columns on some
            # platforms, but on Windows Tk 8.6 it typically counts them as 1.
            # However, Python's len() counts them as 1 too, so the real issue
            # is that some emoji are represented as surrogate pairs in UTF-16
            # which Tcl/Tk internally uses. We need to map Python offsets to
            # Tcl/Tk character indices.
            def py_offset_to_tk(offset, _line=line):
                """Convert a Python string index to a Tkinter text widget column index.

                Tcl/Tk 8.6 on Windows uses UTF-16 internally, so characters
                outside the BMP (code point > 0xFFFF) occupy 2 Tk indices.
                """
                tk_col = 0
                for i in range(min(offset, len(_line))):
                    cp = ord(_line[i])
                    if cp > 0xFFFF:
                        tk_col += 2  # surrogate pair in UTF-16
                    else:
                        tk_col += 1
                return tk_col

            def add_tag(pattern, tag, _line=line, _row=row, flags=0):
                for m in re.finditer(pattern, _line, flags):
                    start = m.start(1) if m.lastindex else m.start()
                    end = m.end(1) if m.lastindex else m.end()
                    tk_start = py_offset_to_tk(start, _line)
                    tk_end = py_offset_to_tk(end, _line)
                    self.editor.tag_add(tag, f"{_row}.{tk_start}", f"{_row}.{tk_end}")

            if re.match(r"^\s*---\s*$", line):
                self.editor.tag_add("hl_separator", f"{row}.0", f"{row}.end")
                continue

            stripped = line.strip()

            # comments get full-line treatment outside SVG
            if not in_svg_block and (
                re.match(r"\s*//", line) or re.match(r"\s*#", line)
            ):
                self.editor.tag_add("hl_comment", f"{row}.0", f"{row}.end")
                continue

            # SVG start line
            if not in_svg_block and re.match(r"\s*svg(\s|$)", line):
                add_tag(string_pattern, "hl_string")
                add_tag(pipe_pattern, "hl_pipe")
                add_tag(shape_pattern, "hl_shape")
                add_tag(color_alias_pattern, "hl_color")
                add_tag(color_variant_pattern, "hl_color")
                add_tag(hex_color_pattern, "hl_color")
                add_tag(key_pattern, "hl_key")
                add_tag(number_pattern, "hl_value")
                add_tag(equals_pattern, "hl_equals")
                in_svg_block = True
                continue

            # Inside SVG block
            if in_svg_block:
                if re.match(endsvg_pattern, line):
                    self.editor.tag_add("hl_svg_endsvg", f"{row}.0", f"{row}.end")
                    in_svg_block = False
                    continue

                add_tag(svg_comment_pattern, "hl_svg_comment")
                add_tag(svg_tag_pattern, "hl_svg_tag")
                add_tag(svg_attr_pattern, "hl_svg_attr")
                add_tag(string_pattern, "hl_string")
                add_tag(hex_color_pattern, "hl_color")
                add_tag(equals_pattern, "hl_equals")
                continue

            # table sub-lines
            if re.match(r"\s*(cols|header|row)=", line):
                add_tag(r"^(cols|header|row)(?==)", "hl_key")
                add_tag(string_pattern, "hl_string")
                add_tag(number_pattern, "hl_value")
                add_tag(equals_pattern, "hl_equals")
                continue

            # normal shape lines
            add_tag(string_pattern, "hl_string")
            add_tag(pipe_pattern, "hl_pipe")
            add_tag(shape_pattern, "hl_shape")
            add_tag(color_alias_pattern, "hl_color")
            add_tag(color_variant_pattern, "hl_color")
            add_tag(hex_color_pattern, "hl_color")
            add_tag(key_pattern, "hl_key")
            add_tag(number_pattern, "hl_value")
            add_tag(equals_pattern, "hl_equals")

    # ── helpers ───────────────────────────────────────────────────────────────
    def _set_status(self, msg, level="info"):
        colors = {
            "info": "#B8B8B8",
            "success": "#70AD47",  # Office green
            "error": "#C55A5A",
            "warning": "#FFC000",  # Office gold
        }
        self.status.config(text=msg, fg=colors.get(level, "#B8B8B8"))

    def _get_shapes(self):
        """Return a flat list of shapes from the first (or only) slide block."""
        slides = self._get_slides()
        return slides[0] if slides else []

    def _get_slides(self):
        """
        Return list-of-lists: one inner list of shapes per slide block.
        Refreshes theme colors and calls parse_dsl_slides.
        """
        text = self.editor.get("1.0", "end").strip()
        if not text:
            return [[]]
        refresh_dsl_theme_colors()
        return parse_dsl_slides(text)

    def _lock(self):
        self.copy_btn.config(state="disabled", text="⏳  Working…")
        self.insert_btn.config(state="disabled", text="⏳  Working…")
        self.slide_btn.config(state="disabled", text="⏳  Working…")

    def _unlock(self):
        self.copy_btn.config(state="normal", text="📋  Copy to Clipboard")
        self.insert_btn.config(state="normal", text="➕  Insert on Current Slide")
        self.slide_btn.config(state="normal", text="📄  Create Full Slide")

    # ── copy to clipboard ─────────────────────────────────────────────────────
    def copy_to_clipboard(self):
        try:
            shapes = self._get_shapes()
        except Exception as e:
            messagebox.showerror("Parse Error", str(e))
            return

        if not shapes:
            messagebox.showwarning("Empty", "Nothing to copy.")
            return

        self._lock()

        def run():
            try:
                self.connector.create_shapes_and_copy(
                    shapes,
                    status_cb=lambda m: self.root.after(0, self._set_status, m, "info"),
                )
                self.root.after(0, self._after_copy)
            except Exception as e:
                self.root.after(0, self._on_error, str(e))

        threading.Thread(target=run, daemon=True).start()

    def _after_copy(self):
        self._unlock()
        self._set_status(
            "✓ Shapes copied — switch to PowerPoint and press Ctrl+V", "success"
        )

    # ── create full slide ─────────────────────────────────────────────────────
    def create_full_slide(self):
        try:
            slides_data = self._get_slides()
        except Exception as e:
            messagebox.showerror("Parse Error", str(e))
            return

        total_shapes = sum(len(s) for s in slides_data)
        if total_shapes == 0:
            messagebox.showwarning("Empty", "Nothing to create.")
            return

        self._lock()

        def run():
            try:
                self.connector.create_on_new_slide(
                    slides_data,
                    status_cb=lambda m: self.root.after(0, self._set_status, m, "info"),
                )
                self.root.after(0, self._after_slide)
            except Exception as e:
                self.root.after(0, self._on_error, str(e))

        threading.Thread(target=run, daemon=True).start()

    def insert_on_current_slide(self):
        """Insert shapes on the active slide without creating a new one"""

        try:
            shapes = self._get_shapes()
        except Exception as e:
            messagebox.showerror("Parse Error", str(e))
            return

        if not shapes:
            messagebox.showwarning("Empty", "Nothing to insert.")
            return

        self._lock()

        def run():
            try:
                self.connector.create_on_current_slide(
                    shapes,
                    status_cb=lambda m: self.root.after(0, self._set_status, m, "info"),
                )
                self.root.after(0, self._after_insert)
            except Exception as e:
                self.root.after(0, self._on_error, str(e))

        threading.Thread(target=run, daemon=True).start()

    def _after_insert(self):
        self._unlock()
        self._set_status("✓ Shapes inserted on current slide!", "success")

    def _after_slide(self):
        self._unlock()
        # Detailed count already shown via status_cb during creation
        self._set_status("✓ Slide(s) created successfully!", "success")

    # ── error handler ─────────────────────────────────────────────────────────
    def _on_error(self, msg):
        self._unlock()
        self._set_status(f"Error: {msg}", "error")

        if "win32com" in msg or "No module" in msg:
            messagebox.showerror(
                "Missing Dependency",
                "pywin32 is required.\n\nRun:  pip install pywin32\n\n"
                "Microsoft PowerPoint must also be installed.",
            )
        else:
            messagebox.showerror("Error", msg)

    def run(self):
        self.root.mainloop()


# ── entry point ───────────────────────────────────────────────────────────────
if __name__ == "__main__":
    App().run()
