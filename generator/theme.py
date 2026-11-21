"""Theme system for presentation generation.

Defines color palettes, typography presets, slide layout dimensions,
and pre-built themes (BW_DOODLE, DARK_MODERN, EXECUTIVE_LIGHT).
"""

from __future__ import annotations

from dataclasses import dataclass, field
from typing import Dict, Any
from pptx.util import Pt, Inches, Emu
from pptx.dml.color import RGBColor


@dataclass
class ColorPalette:
    """Color palette definition storing hex color values."""
    background: str       # Slide background
    surface: str          # Card/container backgrounds
    primary: str          # Main accent color
    secondary: str        # Secondary accent
    text_primary: str     # Main text color
    text_secondary: str   # Muted text color
    text_on_primary: str  # Text on primary-colored bg
    border: str           # Card borders
    success: str          # Green accent
    warning: str          # Yellow/orange accent
    danger: str           # Red accent

    def rgb(self, color_name: str) -> RGBColor:
        """Get RGBColor from hex string of named color attribute."""
        if not hasattr(self, color_name):
            raise AttributeError(f"ColorPalette has no color named '{color_name}'")
        hex_val = getattr(self, color_name).lstrip('#')
        if len(hex_val) == 3:
            hex_val = ''.join(c * 2 for c in hex_val)
        return RGBColor(int(hex_val[0:2], 16), int(hex_val[2:4], 16), int(hex_val[4:6], 16))

    @staticmethod
    def hex_to_rgb(hex_color: str) -> RGBColor:
        """Convert a hex color string (e.g. '#1A1A1A' or '1A1A1A') to an RGBColor."""
        hex_val = hex_color.lstrip('#')
        if len(hex_val) == 3:
            hex_val = ''.join(c * 2 for c in hex_val)
        if len(hex_val) != 6:
            raise ValueError(f"Invalid hex color string: '{hex_color}'")
        return RGBColor(int(hex_val[0:2], 16), int(hex_val[2:4], 16), int(hex_val[4:6], 16))


@dataclass
class Typography:
    """Typography settings and font size presets in Points."""
    heading_font: str = "Calibri"
    body_font: str = "Calibri"
    mono_font: str = "Consolas"
    # Size presets (Points)
    title_size: int = 36        # Main slide title
    heading_size: int = 24      # Section header / card title
    subheading_size: int = 16   # Subtitles and section labels
    body_size: int = 12         # Standard body text
    caption_size: int = 10      # Footnotes and captions
    stat_size: int = 48         # Big metric callouts
    badge_size: int = 9         # Pill tags and badges


@dataclass
class SlideLayout:
    """Coordinates and dimensions for 16:9 widescreen slide elements in Inches."""
    slide_width: float = 13.333
    slide_height: float = 7.5
    margin_left: float = 0.7
    margin_right: float = 0.7
    margin_top: float = 0.6
    margin_bottom: float = 0.5
    card_padding: float = 0.25
    card_radius: float = 0.15   # Corner radius for cards
    card_gap: float = 0.3       # Gap between cards

    @property
    def content_width(self) -> float:
        """Available horizontal content width inside margins in Inches."""
        return self.slide_width - self.margin_left - self.margin_right

    @property
    def content_height(self) -> float:
        """Available vertical content height inside margins in Inches."""
        return self.slide_height - self.margin_top - self.margin_bottom


@dataclass
class Theme:
    """Presentation theme encapsulating colors, typography, and slide layout."""
    name: str
    colors: ColorPalette
    typography: Typography = field(default_factory=Typography)
    layout: SlideLayout = field(default_factory=SlideLayout)


# ==============================================================================
# Pre-built Themes
# ==============================================================================

# 1. BW_DOODLE - Black-and-white character explainer style
BW_DOODLE = Theme(
    name="BW_DOODLE",
    colors=ColorPalette(
        background="#FDFBF7",      # Warm off-white / cream
        surface="#FFFFFF",         # Card background
        primary="#1A1A1A",         # Deep charcoal / black
        secondary="#4A4A4A",       # Medium charcoal
        text_primary="#1A1A1A",    # Dark text
        text_secondary="#6B6B6B",  # Muted grey text
        text_on_primary="#FFFFFF", # White text on dark primary
        border="#E0E0E0",          # Light border
        success="#2D8A4E",         # Green accent
        warning="#D4A843",         # Warm amber accent
        danger="#C0392B",          # Crimson accent
    ),
    typography=Typography(
        heading_font="Poppins",
        body_font="Inter",
        mono_font="Consolas",
    ),
    layout=SlideLayout(),
)

# 2. DARK_MODERN - Sleek modern dark tech style
DARK_MODERN = Theme(
    name="DARK_MODERN",
    colors=ColorPalette(
        background="#0F172A",      # Slate 900
        surface="#1E293B",         # Slate 800
        primary="#3B82F6",         # Blue 500
        secondary="#8B5CF6",       # Violet 500
        text_primary="#F8FAFC",    # Slate 50
        text_secondary="#94A3B8",  # Slate 400
        text_on_primary="#FFFFFF", # White text
        border="#334155",          # Slate 700
        success="#10B981",         # Emerald 500
        warning="#F59E0B",         # Amber 500
        danger="#EF4444",          # Red 500
    ),
    typography=Typography(
        heading_font="Segoe UI",
        body_font="Segoe UI",
        mono_font="Consolas",
    ),
    layout=SlideLayout(),
)

# 3. EXECUTIVE_LIGHT - Clean professional corporate style
EXECUTIVE_LIGHT = Theme(
    name="EXECUTIVE_LIGHT",
    colors=ColorPalette(
        background="#FFFFFF",      # Pure white
        surface="#F8FAFC",         # Crisp cool light grey
        primary="#2563EB",         # Royal blue 600
        secondary="#7C3AED",       # Deep purple 600
        text_primary="#1E293B",    # Dark slate text
        text_secondary="#64748B",  # Cool grey text
        text_on_primary="#FFFFFF", # White text
        border="#E2E8F0",          # Subtle border
        success="#059669",         # Emerald green 600
        warning="#D97706",         # Warm amber 600
        danger="#DC2626",          # Vibrant red 600
    ),
    typography=Typography(
        heading_font="Calibri",
        body_font="Calibri",
        mono_font="Consolas",
    ),
    layout=SlideLayout(),
)

THEMES: dict[str, Theme] = {
    "BW_DOODLE": BW_DOODLE,
    "DARK_MODERN": DARK_MODERN,
    "EXECUTIVE_LIGHT": EXECUTIVE_LIGHT,
}

_THEME_ALIASES: dict[str, str] = {
    "BW_DOODLE": "BW_DOODLE",
    "BW": "BW_DOODLE",
    "DOODLE": "BW_DOODLE",
    "BW-DOODLE": "BW_DOODLE",
    "DARK_MODERN": "DARK_MODERN",
    "DARK": "DARK_MODERN",
    "DARK-MODERN": "DARK_MODERN",
    "EXECUTIVE_LIGHT": "EXECUTIVE_LIGHT",
    "EXECUTIVE": "EXECUTIVE_LIGHT",
    "LIGHT": "EXECUTIVE_LIGHT",
    "EXECUTIVE-LIGHT": "EXECUTIVE_LIGHT",
}


def get_theme(name: str) -> Theme:
    """Retrieve a Theme instance by name or alias (case-insensitive).

    Args:
        name: Name of the theme (e.g., 'BW_DOODLE', 'dark_modern', 'executive_light').

    Returns:
        The corresponding Theme instance.

    Raises:
        ValueError: If the theme name is not recognized.
    """
    normalized = name.strip().upper().replace(" ", "_")
    theme_key = _THEME_ALIASES.get(normalized, normalized)

    if theme_key in THEMES:
        return THEMES[theme_key]

    valid = ", ".join(repr(k) for k in THEMES.keys())
    raise ValueError(f"Unknown theme '{name}'. Available themes: {valid}")
