"""
Theme configuration and styling tokens for the Unified Copilot Suite.
"""

from dataclasses import dataclass
from pathlib import Path
import os


@dataclass(frozen=True)
class AppPalette:
    # Backgrounds (Clean Neutral Dark)
    BG_DARK = "#09090B"
    BG_SURFACE = "#121215"
    BG_CARD = "#161619"
    BG_CARD_HOVER = "#202024"
    BG_INPUT = "#0D0D10"
    BORDER_COLOR = "#27272A"
    BORDER_LIGHT = "#3F3F46"

    # Text (Clean Neutral)
    TEXT_PRIMARY = "#F4F4F5"
    TEXT_SECONDARY = "#A1A1AA"
    TEXT_MUTED = "#71717A"

    # Brand Colors
    BRAND_PPT = "#D24726"      # PowerPoint Red-Orange
    BRAND_WORD = "#2B579A"     # Word Classic Blue
    BRAND_EXCEL = "#217346"    # Excel Emerald Green
    BRAND_CV = "#7C3AED"       # CV Builder Royal Violet
    BRAND_FILES = "#D946EF"  # Files Copilot Magenta

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
