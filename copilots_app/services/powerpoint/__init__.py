"""PowerPoint Copilot service facade."""

from copilots_app.services.powerpoint.constants import VALID_SHAPE_TYPES, DEFAULT_DSL_COLORS, THEME_COLORS
from copilots_app.services.powerpoint.parser import parse_dsl, parse_dsl_slides, refresh_dsl_theme_colors
from copilots_app.services.powerpoint.connector import PowerPointConnector
