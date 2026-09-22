"""PowerPoint Copilot service facade."""

from copilots_app.services.powerpoint.constants import VALID_SHAPE_TYPES, DEFAULT_DSL_COLORS, THEME_COLORS
from copilots_app.services.powerpoint.parser import parse_dsl, parse_dsl_slides, refresh_dsl_theme_colors
from copilots_app.services.powerpoint.connector import PowerPointConnector
from copilots_app.services.powerpoint.extractor import PowerPointExtractor
from copilots_app.services.powerpoint.edit_parser import parse_edit_dsl, looks_like_edit_dsl
from copilots_app.services.powerpoint.editor import PowerPointEditor
