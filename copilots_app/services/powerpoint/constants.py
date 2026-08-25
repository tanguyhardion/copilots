"""
PowerPoint constants, shape registry, and theme definitions.
"""

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
