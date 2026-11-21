"""
slide_builder.py - High-Level Visual Component API for PPTX Generation

This module provides the creative API that the LLM calls to build presentation slides.
Instead of raw python-pptx coordinate math, the LLM calls expressive methods like:
    slide.add_hero_heading("Title", badge="OVERVIEW")
    slide.add_character("thinking", side="left", speech="Hmm...")
    slide.add_bento_cards([...])

The slide builder handles all layout math, text wrapping, element spacing,
and visual styling automatically.
"""

from __future__ import annotations

import os
from pathlib import Path
from typing import Optional
from dataclasses import dataclass

from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.chart import XL_CHART_TYPE

from .theme import Theme, get_theme, ColorPalette
from .primitives import (
    set_slide_background,
    add_rounded_rect,
    add_text_box,
    add_styled_text_box,
    add_image,
    add_line,
    add_badge,
    add_accent_bar,
)

# ---------------------------------------------------------------------------
# Paths
# ---------------------------------------------------------------------------
PROJECT_ROOT = Path(__file__).resolve().parent.parent
ASSETS_DIR = PROJECT_ROOT / "assets"
CHARACTERS_DIR = ASSETS_DIR / "characters"
ICONS_DIR = ASSETS_DIR / "icons"


# ---------------------------------------------------------------------------
# SlideCanvas - the per-slide creative API
# ---------------------------------------------------------------------------
class SlideCanvas:
    """High-level API for composing a single presentation slide.

    Every ``add_*`` method places visual elements at automatically
    computed positions so they never overlap.  The caller (the LLM-generated
    Python script) only describes *what* should appear, not *where* each
    pixel goes.
    """

    def __init__(self, slide, theme: Theme, prs: Presentation):
        self.slide = slide
        self.theme = theme
        self.prs = prs
        self._cursor_y: float = theme.layout.margin_top  # inches from top
        self._used_regions: list[tuple] = []  # track placed elements

        # Paint the background
        set_slide_background(self.slide, theme.colors.hex_to_rgb(theme.colors.background))

    # -- helpers -----------------------------------------------------------
    def _advance_cursor(self, height: float, gap: float = 0.3) -> float:
        """Move the vertical cursor down and return the starting Y."""
        y = self._cursor_y
        self._cursor_y += height + gap
        return y

    @property
    def _left(self) -> float:
        return self.theme.layout.margin_left

    @property
    def _cw(self) -> float:
        return self.theme.layout.content_width

    @property
    def _ch(self) -> float:
        return self.theme.layout.content_height

    def _icon_path(self, name: str) -> Optional[str]:
        """Resolve an icon name to its SVG path."""
        # Try exact match first, then with .svg extension
        for suffix in ["", ".svg"]:
            p = ICONS_DIR / f"{name}{suffix}"
            if p.exists():
                return str(p)
        # Try replacing hyphens with underscores
        for suffix in ["", ".svg"]:
            p = ICONS_DIR / f"{name.replace('-', '_')}{suffix}"
            if p.exists():
                return str(p)
        return None

    def _character_path(self, pose: str) -> Optional[str]:
        """Resolve a character pose name to its SVG path."""
        for suffix in ["", ".svg"]:
            p = CHARACTERS_DIR / f"{pose}{suffix}"
            if p.exists():
                return str(p)
        return None

    # =====================================================================
    #  PUBLIC API - Slide Components
    # =====================================================================

    # ----- 1. TITLE / HERO SLIDE -----------------------------------------
    def add_hero_heading(
        self,
        title: str,
        subtitle: str = "",
        badge: str = "",
        author: str = "",
    ):
        """Full-width title slide with large heading, optional subtitle and badge.

        Args:
            title: Main heading text (displayed large and bold).
            subtitle: Smaller descriptive line beneath the title.
            badge: Small pill label above the title (e.g., 'EXECUTIVE SUMMARY').
            author: Attribution line at the bottom.
        """
        colors = self.theme.colors
        typo = self.theme.typography
        left = self._left
        cw = self._cw

        y = self._advance_cursor(0)  # start at cursor

        # Badge pill
        if badge:
            add_badge(
                self.slide, left, y, badge.upper(),
                bg_color=colors.hex_to_rgb(colors.primary),
                text_color=colors.hex_to_rgb(colors.text_on_primary),
                font_size=typo.badge_size,
            )
            y += 0.45

        # Title
        add_text_box(
            self.slide, left, y, cw, 1.6,
            text=title,
            font_name=typo.heading_font,
            font_size=typo.title_size,
            font_color=colors.hex_to_rgb(colors.text_primary),
            bold=True,
            alignment=PP_ALIGN.LEFT,
            vertical_anchor=MSO_ANCHOR.MIDDLE,
        )
        y += 1.7

        # Subtitle
        if subtitle:
            add_text_box(
                self.slide, left, y, cw * 0.7, 0.8,
                text=subtitle,
                font_name=typo.body_font,
                font_size=typo.subheading_size,
                font_color=colors.hex_to_rgb(colors.text_secondary),
                alignment=PP_ALIGN.LEFT,
            )
            y += 0.9

        # Accent line
        add_line(
            self.slide,
            left, y, left + 2.0, y,
            color=colors.hex_to_rgb(colors.primary),
            width=3.0,
        )
        y += 0.4

        # Author
        if author:
            add_text_box(
                self.slide, left, y, cw, 0.4,
                text=author,
                font_name=typo.body_font,
                font_size=typo.caption_size,
                font_color=colors.hex_to_rgb(colors.text_secondary),
                alignment=PP_ALIGN.LEFT,
            )
            y += 0.5

        self._cursor_y = y

    # ----- 2. SECTION HEADER --------------------------------------------
    def add_section_header(self, title: str, category: str = ""):
        """Section divider with category badge and large title."""
        colors = self.theme.colors
        typo = self.theme.typography
        left = self._left
        cw = self._cw

        y = self._advance_cursor(0)

        if category:
            add_badge(
                self.slide, left, y, category.upper(),
                bg_color=colors.hex_to_rgb(colors.primary),
                text_color=colors.hex_to_rgb(colors.text_on_primary),
                font_size=typo.badge_size,
            )
            y += 0.45

        add_text_box(
            self.slide, left, y, cw, 1.0,
            text=title,
            font_name=typo.heading_font,
            font_size=typo.heading_size,
            font_color=colors.hex_to_rgb(colors.text_primary),
            bold=True,
            alignment=PP_ALIGN.LEFT,
        )
        self._cursor_y = y + 1.2

    # ----- 3. CHARACTER POSE --------------------------------------------
    def add_character(
        self,
        pose: str = "thinking",
        side: str = "left",
        speech_bubble: str = "",
        size: float = 2.5,
    ):
        """Place a doodle character on the slide.

        Args:
            pose: Character pose name (thinking, pointing, celebrating,
                  confused, presenting, reading).
            side: 'left' or 'right' - which side of the slide.
            speech_bubble: Optional text for a speech bubble near the character.
            size: Width/height of the character in inches.
        """
        colors = self.theme.colors
        typo = self.theme.typography
        layout = self.theme.layout

        char_path = self._character_path(pose)

        if side == "left":
            char_x = layout.margin_left
        else:
            char_x = layout.slide_width - layout.margin_right - size

        char_y = self._cursor_y

        if char_path:
            add_image(self.slide, char_path, char_x, char_y, size, size)
        else:
            # Fallback: draw a simple circle-head character placeholder
            head_size = size * 0.35
            head_x = char_x + (size - head_size) / 2
            head_y = char_y + size * 0.05
            shape = self.slide.shapes.add_shape(
                MSO_SHAPE.OVAL, Inches(head_x), Inches(head_y),
                Inches(head_size), Inches(head_size),
            )
            shape.fill.solid()
            shape.fill.fore_color.rgb = RGBColor(0x1A, 0x1A, 0x1A)
            shape.line.fill.background()

            # Body line
            body = self.slide.shapes.add_shape(
                MSO_SHAPE.ROUNDED_RECTANGLE,
                Inches(char_x + size * 0.3), Inches(char_y + size * 0.42),
                Inches(size * 0.4), Inches(size * 0.5),
            )
            body.fill.solid()
            body.fill.fore_color.rgb = RGBColor(0x1A, 0x1A, 0x1A)
            body.line.fill.background()

        # Speech bubble
        if speech_bubble:
            if side == "left":
                bubble_x = char_x + size + 0.3
                bubble_w = self._cw - size - 0.6
            else:
                bubble_x = layout.margin_left
                bubble_w = self._cw - size - 0.6

            bubble_y = char_y + 0.2
            bubble_h = min(1.5, max(0.6, len(speech_bubble) / 60 * 0.5 + 0.5))

            # Bubble background
            add_rounded_rect(
                self.slide, bubble_x, bubble_y, bubble_w, bubble_h,
                fill_color=colors.hex_to_rgb(colors.surface),
                border_color=colors.hex_to_rgb(colors.border),
                border_width=1.0,
                corner_radius=0.1,
            )
            # Bubble text
            add_text_box(
                self.slide, bubble_x + 0.2, bubble_y + 0.1,
                bubble_w - 0.4, bubble_h - 0.2,
                text=speech_bubble,
                font_name=typo.body_font,
                font_size=typo.body_size + 1,
                font_color=colors.hex_to_rgb(colors.text_primary),
                italic=True,
                alignment=PP_ALIGN.LEFT,
            )

        self._cursor_y = char_y + size + 0.2

    # ----- 4. BENTO GRID CARDS ------------------------------------------
    def add_bento_cards(
        self,
        cards: list[dict],
        columns: int = 0,
        card_height: float = 0,
    ):
        """Grid of styled content cards.

        Args:
            cards: List of dicts, each with keys:
                - title (str): Card heading.
                - body (str): Card description text.
                - icon (str, optional): Icon name from assets/icons/.
                - highlight (str, optional): Bold stat or metric callout.
            columns: Number of columns (auto-detected from card count if 0).
            card_height: Height of each card in inches (auto if 0).
        """
        colors = self.theme.colors
        typo = self.theme.typography
        layout = self.theme.layout
        left = layout.margin_left
        cw = self._cw
        gap = layout.card_gap

        n = len(cards)
        if columns == 0:
            columns = min(n, 3) if n <= 6 else 4

        rows = (n + columns - 1) // columns
        card_w = (cw - gap * (columns - 1)) / columns
        if card_height == 0:
            card_height = min(2.8, max(1.8, 4.5 / rows))

        y = self._advance_cursor(0)

        for idx, card in enumerate(cards):
            row = idx // columns
            col = idx % columns
            cx = left + col * (card_w + gap)
            cy = y + row * (card_height + gap)

            # Card background
            add_rounded_rect(
                self.slide, cx, cy, card_w, card_height,
                fill_color=colors.hex_to_rgb(colors.surface),
                border_color=colors.hex_to_rgb(colors.border),
                border_width=0.75,
                corner_radius=0.1,
            )

            # Accent bar on left edge
            add_accent_bar(
                self.slide, cx + 0.08, cy + 0.15,
                card_height - 0.3,
                color=colors.hex_to_rgb(colors.primary),
                width=0.04,
            )

            inner_x = cx + 0.25
            inner_w = card_w - 0.45
            content_y = cy + 0.15

            # Icon
            icon_name = card.get("icon", "")
            icon_path = self._icon_path(icon_name) if icon_name else None
            if icon_path:
                add_image(self.slide, icon_path, inner_x, content_y, 0.35, 0.35)
                content_y += 0.45
            else:
                content_y += 0.05

            # Title
            title = card.get("title", "")
            if title:
                add_text_box(
                    self.slide, inner_x, content_y, inner_w, 0.4,
                    text=title,
                    font_name=typo.heading_font,
                    font_size=typo.subheading_size - 2,
                    font_color=colors.hex_to_rgb(colors.text_primary),
                    bold=True,
                    alignment=PP_ALIGN.LEFT,
                )
                content_y += 0.42

            # Highlight/stat
            highlight = card.get("highlight", "")
            if highlight:
                add_text_box(
                    self.slide, inner_x, content_y, inner_w, 0.35,
                    text=highlight,
                    font_name=typo.heading_font,
                    font_size=typo.heading_size,
                    font_color=colors.hex_to_rgb(colors.primary),
                    bold=True,
                    alignment=PP_ALIGN.LEFT,
                )
                content_y += 0.4

            # Body text
            body = card.get("body", "")
            if body:
                remaining = card_height - (content_y - cy) - 0.15
                add_text_box(
                    self.slide, inner_x, content_y, inner_w, max(0.4, remaining),
                    text=body,
                    font_name=typo.body_font,
                    font_size=typo.body_size,
                    font_color=colors.hex_to_rgb(colors.text_secondary),
                    alignment=PP_ALIGN.LEFT,
                )

        self._cursor_y = y + rows * (card_height + gap)

    # ----- 5. STAT BANNER -----------------------------------------------
    def add_stat_banner(self, metrics: list[dict]):
        """Horizontal row of big-number statistics.

        Args:
            metrics: List of dicts with keys:
                - value (str): The big number/stat (e.g., '10x', '99.9%', '$2.4M').
                - label (str): Small description beneath the number.
                - icon (str, optional): Icon name.
        """
        colors = self.theme.colors
        typo = self.theme.typography
        left = self._left
        cw = self._cw
        n = len(metrics)
        gap = 0.3

        metric_w = (cw - gap * (n - 1)) / n
        metric_h = 1.6
        y = self._advance_cursor(0)

        for i, m in enumerate(metrics):
            mx = left + i * (metric_w + gap)

            # Card bg
            add_rounded_rect(
                self.slide, mx, y, metric_w, metric_h,
                fill_color=colors.hex_to_rgb(colors.surface),
                border_color=colors.hex_to_rgb(colors.border),
                border_width=0.75,
                corner_radius=0.1,
            )

            # Big stat value
            add_text_box(
                self.slide, mx + 0.2, y + 0.15, metric_w - 0.4, 0.8,
                text=m.get("value", ""),
                font_name=typo.heading_font,
                font_size=typo.stat_size,
                font_color=colors.hex_to_rgb(colors.primary),
                bold=True,
                alignment=PP_ALIGN.CENTER,
                vertical_anchor=MSO_ANCHOR.MIDDLE,
            )

            # Label
            add_text_box(
                self.slide, mx + 0.2, y + 1.0, metric_w - 0.4, 0.4,
                text=m.get("label", ""),
                font_name=typo.body_font,
                font_size=typo.body_size,
                font_color=colors.hex_to_rgb(colors.text_secondary),
                alignment=PP_ALIGN.CENTER,
            )

        self._cursor_y = y + metric_h + 0.3

    # ----- 6. BULLET LIST -----------------------------------------------
    def add_bullet_list(
        self,
        items: list[str],
        title: str = "",
        position: str = "full",
        icon: str = "",
    ):
        """Styled bullet-point list.

        Args:
            items: List of text strings for each bullet point.
            title: Optional heading above the list.
            position: 'full' (full width), 'left' (left half), 'right' (right half).
            icon: Optional icon for the list header.
        """
        colors = self.theme.colors
        typo = self.theme.typography
        left = self._left
        cw = self._cw

        if position == "left":
            w = cw * 0.47
        elif position == "right":
            left = left + cw * 0.53
            w = cw * 0.47
        else:
            w = cw

        y = self._advance_cursor(0)

        if title:
            add_text_box(
                self.slide, left, y, w, 0.5,
                text=title,
                font_name=typo.heading_font,
                font_size=typo.subheading_size,
                font_color=colors.hex_to_rgb(colors.text_primary),
                bold=True,
            )
            y += 0.55

        for item in items:
            # Bullet dot
            dot_shape = self.slide.shapes.add_shape(
                MSO_SHAPE.OVAL,
                Inches(left + 0.1), Inches(y + 0.12),
                Inches(0.1), Inches(0.1),
            )
            dot_shape.fill.solid()
            dot_shape.fill.fore_color.rgb = colors.hex_to_rgb(colors.primary)
            dot_shape.line.fill.background()

            # Bullet text
            item_h = max(0.35, len(item) / 90 * 0.4 + 0.25)
            add_text_box(
                self.slide, left + 0.35, y, w - 0.45, item_h,
                text=item,
                font_name=typo.body_font,
                font_size=typo.body_size + 1,
                font_color=colors.hex_to_rgb(colors.text_primary),
            )
            y += item_h + 0.08

        self._cursor_y = y + 0.1

    # ----- 7. COMPARISON (TWO-COLUMN) ------------------------------------
    def add_comparison(
        self,
        left_data: dict,
        right_data: dict,
    ):
        """Side-by-side comparison layout.

        Args:
            left_data: Dict with 'title', 'points' (list[str]), 'color' (optional hex).
            right_data: Same structure as left_data.
        """
        colors = self.theme.colors
        typo = self.theme.typography
        left = self._left
        cw = self._cw
        gap = 0.4
        col_w = (cw - gap) / 2

        y = self._advance_cursor(0)

        for i, data in enumerate([left_data, right_data]):
            cx = left + i * (col_w + gap)
            accent = data.get("color", colors.primary if i == 0 else colors.secondary)

            # Title
            add_text_box(
                self.slide, cx, y, col_w, 0.5,
                text=data.get("title", ""),
                font_name=typo.heading_font,
                font_size=typo.subheading_size,
                font_color=colors.hex_to_rgb(accent),
                bold=True,
            )

            # Accent line under title
            add_line(
                self.slide, cx, y + 0.5, cx + col_w, y + 0.5,
                color=colors.hex_to_rgb(accent),
                width=2.0,
            )

            # Points
            py = y + 0.7
            for point in data.get("points", []):
                dot = self.slide.shapes.add_shape(
                    MSO_SHAPE.OVAL,
                    Inches(cx + 0.1), Inches(py + 0.1),
                    Inches(0.08), Inches(0.08),
                )
                dot.fill.solid()
                dot.fill.fore_color.rgb = colors.hex_to_rgb(accent)
                dot.line.fill.background()

                ph = max(0.3, len(point) / 80 * 0.35 + 0.2)
                add_text_box(
                    self.slide, cx + 0.3, py, col_w - 0.4, ph,
                    text=point,
                    font_name=typo.body_font,
                    font_size=typo.body_size,
                    font_color=colors.hex_to_rgb(colors.text_primary),
                )
                py += ph + 0.05

        self._cursor_y = y + 4.0

    # ----- 8. TIMELINE / PROCESS STEPS -----------------------------------
    def add_timeline(self, steps: list[dict]):
        """Horizontal process timeline with numbered steps.

        Args:
            steps: List of dicts with keys:
                - title (str): Step heading.
                - description (str): Brief description.
                - icon (str, optional): Icon name.
        """
        colors = self.theme.colors
        typo = self.theme.typography
        left = self._left
        cw = self._cw
        n = len(steps)
        gap = 0.2
        step_w = (cw - gap * (n - 1)) / n

        y = self._advance_cursor(0)

        for i, step in enumerate(steps):
            sx = left + i * (step_w + gap)

            # Step number circle
            circle_size = 0.5
            circle = self.slide.shapes.add_shape(
                MSO_SHAPE.OVAL,
                Inches(sx + (step_w - circle_size) / 2), Inches(y),
                Inches(circle_size), Inches(circle_size),
            )
            circle.fill.solid()
            circle.fill.fore_color.rgb = colors.hex_to_rgb(colors.primary)
            circle.line.fill.background()

            # Step number text
            add_text_box(
                self.slide,
                sx + (step_w - circle_size) / 2, y,
                circle_size, circle_size,
                text=str(i + 1),
                font_name=typo.heading_font,
                font_size=16,
                font_color=colors.hex_to_rgb(colors.text_on_primary),
                bold=True,
                alignment=PP_ALIGN.CENTER,
                vertical_anchor=MSO_ANCHOR.MIDDLE,
            )

            # Connector line to next step
            if i < n - 1:
                line_y = y + circle_size / 2
                add_line(
                    self.slide,
                    sx + (step_w + circle_size) / 2 + 0.05, line_y,
                    sx + step_w + gap - 0.05 + (step_w - circle_size) / 2, line_y,
                    color=colors.hex_to_rgb(colors.border),
                    width=1.5,
                )

            # Step title
            add_text_box(
                self.slide,
                sx, y + 0.65, step_w, 0.4,
                text=step.get("title", ""),
                font_name=typo.heading_font,
                font_size=typo.body_size + 1,
                font_color=colors.hex_to_rgb(colors.text_primary),
                bold=True,
                alignment=PP_ALIGN.CENTER,
            )

            # Step description
            add_text_box(
                self.slide,
                sx, y + 1.05, step_w, 0.8,
                text=step.get("description", ""),
                font_name=typo.body_font,
                font_size=typo.caption_size + 1,
                font_color=colors.hex_to_rgb(colors.text_secondary),
                alignment=PP_ALIGN.CENTER,
            )

        self._cursor_y = y + 2.2

    # ----- 9. QUOTE / CALLOUT -------------------------------------------
    def add_quote(self, text: str, author: str = ""):
        """Large quote callout with optional attribution.

        Args:
            text: The quote text.
            author: Attribution (e.g., 'Elon Musk' or 'From the YouTube lecture').
        """
        colors = self.theme.colors
        typo = self.theme.typography
        left = self._left
        cw = self._cw

        y = self._advance_cursor(0)
        quote_h = max(1.2, len(text) / 80 * 0.5 + 0.8)

        # Background card
        add_rounded_rect(
            self.slide, left, y, cw, quote_h + 0.4,
            fill_color=colors.hex_to_rgb(colors.surface),
            border_color=colors.hex_to_rgb(colors.border),
            border_width=0.5,
            corner_radius=0.12,
        )

        # Left accent bar
        add_accent_bar(
            self.slide, left + 0.15, y + 0.15,
            quote_h + 0.1,
            color=colors.hex_to_rgb(colors.primary),
            width=0.05,
        )

        # Quote mark
        add_text_box(
            self.slide, left + 0.35, y + 0.05, 0.5, 0.5,
            text="\u201C",
            font_name=typo.heading_font,
            font_size=36,
            font_color=colors.hex_to_rgb(colors.primary),
            bold=True,
        )

        # Quote text
        add_text_box(
            self.slide, left + 0.5, y + 0.3, cw - 1.0, quote_h - 0.3,
            text=text,
            font_name=typo.body_font,
            font_size=typo.subheading_size,
            font_color=colors.hex_to_rgb(colors.text_primary),
            italic=True,
            alignment=PP_ALIGN.LEFT,
        )

        # Author
        if author:
            add_text_box(
                self.slide, left + 0.5, y + quote_h, cw - 1.0, 0.3,
                text=f"— {author}",
                font_name=typo.body_font,
                font_size=typo.body_size,
                font_color=colors.hex_to_rgb(colors.text_secondary),
                alignment=PP_ALIGN.RIGHT,
            )

        self._cursor_y = y + quote_h + 0.6

    # ----- 10. NATIVE CHART ---------------------------------------------
    def add_chart(
        self,
        chart_type: str,
        categories: list[str],
        series: list[dict],
        title: str = "",
        height: float = 3.5,
    ):
        """Add a native, editable PowerPoint chart.

        Args:
            chart_type: 'bar', 'column', 'line', 'pie', 'area'.
            categories: List of category labels.
            series: List of dicts with 'name' and 'values' keys.
            title: Optional chart title.
            height: Chart height in inches.
        """
        from pptx.chart.data import CategoryChartData

        colors = self.theme.colors
        typo = self.theme.typography
        left = self._left
        cw = self._cw

        y = self._advance_cursor(0)

        if title:
            add_text_box(
                self.slide, left, y, cw, 0.5,
                text=title,
                font_name=typo.heading_font,
                font_size=typo.subheading_size,
                font_color=colors.hex_to_rgb(colors.text_primary),
                bold=True,
            )
            y += 0.55

        chart_type_map = {
            "bar": XL_CHART_TYPE.BAR_CLUSTERED,
            "column": XL_CHART_TYPE.COLUMN_CLUSTERED,
            "line": XL_CHART_TYPE.LINE,
            "pie": XL_CHART_TYPE.PIE,
            "area": XL_CHART_TYPE.AREA,
        }
        xl_type = chart_type_map.get(chart_type.lower(), XL_CHART_TYPE.COLUMN_CLUSTERED)

        chart_data = CategoryChartData()
        chart_data.categories = categories
        for s in series:
            chart_data.add_series(s["name"], s["values"])

        chart_shape = self.slide.shapes.add_chart(
            xl_type,
            Inches(left), Inches(y),
            Inches(cw), Inches(height),
            chart_data,
        )

        # Style the chart
        chart = chart_shape.chart
        chart.has_legend = len(series) > 1

        # Apply theme colors to series
        palette = [colors.primary, colors.secondary, colors.success, colors.warning, colors.danger]
        for i, s in enumerate(chart.series):
            color_hex = palette[i % len(palette)]
            s.format.fill.solid()
            s.format.fill.fore_color.rgb = colors.hex_to_rgb(color_hex)

        self._cursor_y = y + height + 0.3

    # ----- 11. TABLE ----------------------------------------------------
    def add_table(
        self,
        headers: list[str],
        rows: list[list[str]],
        title: str = "",
    ):
        """Add a styled table.

        Args:
            headers: Column header strings.
            rows: List of rows, each row is a list of cell strings.
            title: Optional table title.
        """
        colors = self.theme.colors
        typo = self.theme.typography
        left = self._left
        cw = self._cw

        y = self._advance_cursor(0)

        if title:
            add_text_box(
                self.slide, left, y, cw, 0.5,
                text=title,
                font_name=typo.heading_font,
                font_size=typo.subheading_size,
                font_color=colors.hex_to_rgb(colors.text_primary),
                bold=True,
            )
            y += 0.55

        num_rows = len(rows) + 1  # +1 for header
        num_cols = len(headers)
        row_h = 0.4
        table_h = num_rows * row_h

        table_shape = self.slide.shapes.add_table(
            num_rows, num_cols,
            Inches(left), Inches(y),
            Inches(cw), Inches(table_h),
        )
        table = table_shape.table

        # Style header row
        for j, h in enumerate(headers):
            cell = table.cell(0, j)
            cell.text = h
            cell.fill.solid()
            cell.fill.fore_color.rgb = colors.hex_to_rgb(colors.primary)
            for paragraph in cell.text_frame.paragraphs:
                for run in paragraph.runs:
                    run.font.color.rgb = colors.hex_to_rgb(colors.text_on_primary)
                    run.font.bold = True
                    run.font.size = Pt(typo.body_size)
                    run.font.name = typo.heading_font

        # Style data rows
        for i, row in enumerate(rows):
            for j, val in enumerate(row):
                cell = table.cell(i + 1, j)
                cell.text = str(val)
                # Alternate row colors
                if i % 2 == 0:
                    cell.fill.solid()
                    cell.fill.fore_color.rgb = colors.hex_to_rgb(colors.surface)
                else:
                    cell.fill.solid()
                    cell.fill.fore_color.rgb = colors.hex_to_rgb(colors.background)

                for paragraph in cell.text_frame.paragraphs:
                    for run in paragraph.runs:
                        run.font.color.rgb = colors.hex_to_rgb(colors.text_primary)
                        run.font.size = Pt(typo.body_size)
                        run.font.name = typo.body_font

        self._cursor_y = y + table_h + 0.3

    # ----- 12. CONTENT BLOCK WITH IMAGE ---------------------------------
    def add_content_with_image(
        self,
        title: str,
        body: str,
        image_path: str = "",
        image_side: str = "right",
    ):
        """Split layout: text on one side, image on the other.

        Args:
            title: Section heading.
            body: Paragraph text.
            image_path: Path to image file. Can also be an icon name.
            image_side: 'left' or 'right'.
        """
        colors = self.theme.colors
        typo = self.theme.typography
        left = self._left
        cw = self._cw
        gap = 0.4

        y = self._advance_cursor(0)
        text_w = cw * 0.55
        img_w = cw * 0.4

        if image_side == "right":
            text_x = left
            img_x = left + text_w + gap
        else:
            img_x = left
            text_x = left + img_w + gap

        # Title
        add_text_box(
            self.slide, text_x, y, text_w, 0.5,
            text=title,
            font_name=typo.heading_font,
            font_size=typo.heading_size - 2,
            font_color=colors.hex_to_rgb(colors.text_primary),
            bold=True,
        )

        # Body
        body_h = max(1.5, len(body) / 60 * 0.4 + 0.5)
        add_text_box(
            self.slide, text_x, y + 0.6, text_w, body_h,
            text=body,
            font_name=typo.body_font,
            font_size=typo.body_size + 1,
            font_color=colors.hex_to_rgb(colors.text_secondary),
        )

        # Image
        resolved = image_path
        if image_path and not os.path.exists(image_path):
            # Try resolving as icon or character name
            resolved = self._icon_path(image_path) or self._character_path(image_path)

        if resolved and os.path.exists(resolved):
            add_image(self.slide, resolved, img_x, y, img_w, img_w * 0.75)
        else:
            # Placeholder rectangle
            add_rounded_rect(
                self.slide, img_x, y, img_w, img_w * 0.75,
                fill_color=colors.hex_to_rgb(colors.surface),
                border_color=colors.hex_to_rgb(colors.border),
            )

        self._cursor_y = y + max(body_h + 0.6, img_w * 0.75) + 0.3

    # ----- 13. FOOTER / PAGE NUMBER -------------------------------------
    def add_footer(self, text: str = "", show_page_number: bool = True, page_number: int = 0):
        """Add a subtle footer bar at the bottom of the slide."""
        colors = self.theme.colors
        typo = self.theme.typography
        layout = self.theme.layout

        footer_y = layout.slide_height - 0.4

        # Thin line
        add_line(
            self.slide,
            layout.margin_left, footer_y - 0.1,
            layout.slide_width - layout.margin_right, footer_y - 0.1,
            color=colors.hex_to_rgb(colors.border),
            width=0.5,
        )

        if text:
            add_text_box(
                self.slide, layout.margin_left, footer_y, 4.0, 0.3,
                text=text,
                font_name=typo.body_font,
                font_size=typo.caption_size,
                font_color=colors.hex_to_rgb(colors.text_secondary),
            )

        if show_page_number and page_number > 0:
            add_text_box(
                self.slide,
                layout.slide_width - layout.margin_right - 1.0, footer_y,
                1.0, 0.3,
                text=str(page_number),
                font_name=typo.body_font,
                font_size=typo.caption_size,
                font_color=colors.hex_to_rgb(colors.text_secondary),
                alignment=PP_ALIGN.RIGHT,
            )


# ---------------------------------------------------------------------------
# PresentationDeck - top-level deck builder
# ---------------------------------------------------------------------------
class PresentationDeck:
    """Top-level presentation builder.

    Usage::

        deck = PresentationDeck(theme="bw_doodle")
        slide = deck.new_slide()
        slide.add_hero_heading("My Title", badge="INTRO")
        slide.add_bento_cards([...])
        deck.save("output.pptx")
    """

    THEME_ALIASES = {
        "bw_doodle": "bw_doodle",
        "dark": "dark_modern",
        "dark_modern": "dark_modern",
        "light": "executive_light",
        "executive": "executive_light",
        "executive_light": "executive_light",
    }

    def __init__(self, theme: str = "bw_doodle"):
        theme_name = self.THEME_ALIASES.get(theme, theme)
        self.theme = get_theme(theme_name)
        self.prs = Presentation()

        # Set 16:9 widescreen
        self.prs.slide_width = Inches(self.theme.layout.slide_width)
        self.prs.slide_height = Inches(self.theme.layout.slide_height)

        self._slide_count = 0

    def new_slide(self) -> SlideCanvas:
        """Create a new blank slide and return its canvas for composition."""
        slide_layout = self.prs.slide_layouts[6]  # Blank layout
        slide = self.prs.slides.add_slide(slide_layout)
        self._slide_count += 1
        return SlideCanvas(slide, self.theme, self.prs)

    @property
    def slide_count(self) -> int:
        return self._slide_count

    def save(self, output_path: str) -> str:
        """Save the presentation to a .pptx file and return the path."""
        os.makedirs(os.path.dirname(os.path.abspath(output_path)), exist_ok=True)
        self.prs.save(output_path)
        return os.path.abspath(output_path)
