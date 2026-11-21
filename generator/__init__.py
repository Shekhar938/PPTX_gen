"""Presentation generation module.

Contains theme definitions, low-level drawing primitives, and slide builders.
"""

from .theme import (
    BW_DOODLE,
    DARK_MODERN,
    EXECUTIVE_LIGHT,
    THEMES,
    ColorPalette,
    SlideLayout,
    Theme,
    Typography,
    get_theme,
)
from .primitives import (
    add_accent_bar,
    add_badge,
    add_image,
    add_line,
    add_rounded_rect,
    add_styled_text_box,
    add_text_box,
    set_slide_background,
)

__all__ = [
    "BW_DOODLE",
    "DARK_MODERN",
    "EXECUTIVE_LIGHT",
    "THEMES",
    "ColorPalette",
    "SlideLayout",
    "Theme",
    "Typography",
    "get_theme",
    "add_accent_bar",
    "add_badge",
    "add_image",
    "add_line",
    "add_rounded_rect",
    "add_styled_text_box",
    "add_text_box",
    "set_slide_background",
]
