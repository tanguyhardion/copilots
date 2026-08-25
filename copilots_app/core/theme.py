"""
Theme configuration, styling tokens, and Qt Stylesheet generator for the Unified Copilot Suite.
"""

from dataclasses import dataclass
from pathlib import Path
import os


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


def get_asset_path(relative_path: str) -> str:
    """Resolve asset path relative to root assets directory."""
    base_dir = Path(__file__).resolve().parent.parent.parent / "assets"
    full_path = base_dir / relative_path
    return str(full_path).replace("\\", "/")


def get_app_stylesheet() -> str:
    """Generate global application QSS styling."""
    return f"""
    * {{
        font-family: "Segoe UI", -apple-system, BlinkMacSystemFont, Roboto, sans-serif;
        color: {AppPalette.TEXT_PRIMARY};
    }}

    QMainWindow, QDialog {{
        background-color: {AppPalette.BG_DARK};
    }}

    QWidget {{
        background-color: transparent;
        font-size: 13px;
    }}

    QScrollBar:vertical {{
        background: {AppPalette.BG_DARK};
        width: 8px;
        margin: 0px;
    }}
    QScrollBar::handle:vertical {{
        background: {AppPalette.BORDER_COLOR};
        min-height: 20px;
        border-radius: 4px;
    }}
    QScrollBar::handle:vertical:hover {{
        background: {AppPalette.BORDER_LIGHT};
    }}
    QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical {{
        height: 0px;
    }}

    QScrollBar:horizontal {{
        background: {AppPalette.BG_DARK};
        height: 8px;
        margin: 0px;
    }}
    QScrollBar::handle:horizontal {{
        background: {AppPalette.BORDER_COLOR};
        min-width: 20px;
        border-radius: 4px;
    }}
    QScrollBar::handle:horizontal:hover {{
        background: {AppPalette.BORDER_LIGHT};
    }}
    QScrollBar::add-line:horizontal, QScrollBar::sub-line:horizontal {{
        width: 0px;
    }}

    QComboBox {{
        background-color: {AppPalette.BG_CARD};
        color: {AppPalette.TEXT_PRIMARY};
        border: 1px solid {AppPalette.BORDER_COLOR};
        border-radius: 6px;
        padding: 5px 10px;
        font-size: 12px;
    }}
    QComboBox:hover {{
        border-color: {AppPalette.BORDER_LIGHT};
    }}
    QComboBox::drop-down {{
        border: none;
        width: 20px;
    }}
    QComboBox QAbstractItemView {{
        background-color: {AppPalette.BG_CARD};
        color: {AppPalette.TEXT_PRIMARY};
        selection-background-color: {AppPalette.PRIMARY};
        selection-color: #FFFFFF;
        border: 1px solid {AppPalette.BORDER_COLOR};
        border-radius: 6px;
        padding: 4px;
        outline: none;
    }}

    QToolTip {{
        background-color: {AppPalette.BG_CARD};
        color: {AppPalette.TEXT_PRIMARY};
        border: 1px solid {AppPalette.BORDER_COLOR};
        border-radius: 4px;
        padding: 4px 8px;
        font-size: 11px;
    }}
    """
