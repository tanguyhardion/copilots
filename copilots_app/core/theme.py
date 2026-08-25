"""
Theme configuration and styling tokens for the Unified Copilot Suite.
"""

from dataclasses import dataclass
import flet as ft


@dataclass(frozen=True)
class AppPalette:
    # Backgrounds
    BG_DARK = "#12141C"
    BG_SURFACE = "#181B26"
    BG_CARD = "#1F2332"
    BG_CARD_HOVER = "#272C3F"
    BG_INPUT = "#161822"
    BORDER_COLOR = "#2D3247"
    BORDER_LIGHT = "#3B425D"

    # Text
    TEXT_PRIMARY = "#F3F4F6"
    TEXT_SECONDARY = "#9CA3AF"
    TEXT_MUTED = "#6B7280"

    # Brand Colors
    BRAND_PPT = "#D24726"      # PowerPoint Red-Orange
    BRAND_WORD = "#2B579A"     # Word Classic Blue
    BRAND_EXCEL = "#217346"    # Excel Emerald Green
    BRAND_CV = "#7C3AED"       # CV Builder Royal Violet

    # Accent and Semantic
    PRIMARY = "#6366F1"        # Indigo
    PRIMARY_HOVER = "#4F46E5"
    SUCCESS = "#10B981"        # Emerald
    WARNING = "#F59E0B"        # Amber
    ERROR = "#EF4444"          # Red
    INFO = "#3B82F6"           # Blue


def create_app_theme() -> ft.Theme:
    """Create the unified Material 3 theme."""
    return ft.Theme(
        color_scheme_seed=AppPalette.PRIMARY,
        visual_density=ft.VisualDensity.COMPACT,
        font_family="Segoe UI, -apple-system, BlinkMacSystemFont, Roboto, sans-serif",
    )

