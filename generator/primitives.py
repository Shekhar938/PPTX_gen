"""Low-level drawing primitives for PPTX slide generation.

Handles shape creation, rounded containers, text frames with typography styling,
vector/raster image embedding with SVG fallbacks, lines, badges, and accent bars.
All coordinate and dimension parameters accept float values representing Inches.
"""

from __future__ import annotations

import io
import os
from typing import Any, List, Dict, Optional, Union

from pptx.presentation import Presentation
from pptx.slide import Slide
from pptx.shapes.autoshape import Shape
from pptx.shapes.picture import Picture
from pptx.shapes.connector import Connector
from pptx.util import Inches, Pt, Emu, Length
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE, MSO_CONNECTOR

from .theme import Theme, ColorPalette


def _to_inches(val: Union[float, int, Length, None]) -> Optional[Length]:
    """Convert a float/int number representing inches to a pptx Length.

    If already a Length, returns it as-is. Returns None if val is None.
    """
    if val is None:
        return None
    if isinstance(val, Length):
        return val
    return Inches(val)


def _ensure_rgb(color: Union[RGBColor, str, None]) -> Optional[RGBColor]:
    """Ensure the color value is an RGBColor instance.

    Converts hex strings (e.g., '#1A1A1A' or 'FFF') using ColorPalette.hex_to_rgb.
    """
    if color is None:
        return None
    if isinstance(color, RGBColor):
        return color
    if isinstance(color, str):
        return ColorPalette.hex_to_rgb(color)
    raise TypeError(f"Expected RGBColor or hex string, got {type(color).__name__}")


def set_slide_background(slide: Slide, color: Union[RGBColor, str]) -> Any:
    """Sets the entire slide background fill color.

    Args:
        slide: The python-pptx Slide object to style.
        color: RGBColor or hex string (e.g. '#FDFBF7') for the background fill.

    Returns:
        The slide background object.
    """
    resolved_color = _ensure_rgb(color)
    background = slide.background
    fill = background.fill
    fill.solid()
    fill.fore_color.rgb = resolved_color
    return background


def add_rounded_rect(
    slide: Slide,
    left: float,
    top: float,
    width: float,
    height: float,
    fill_color: Union[RGBColor, str, None],
    border_color: Union[RGBColor, str, None] = None,
    border_width: float = 0.5,
    corner_radius: float = 0.15,
) -> Shape:
    """Adds a rounded rectangle shape to the slide.

    Uses MSO_SHAPE.ROUNDED_RECTANGLE with calculated corner radius adjustment.

    Args:
        slide: The slide to add the rounded rectangle to.
        left: Left position in inches.
        top: Top position in inches.
        width: Width in inches.
        height: Height in inches.
        fill_color: Fill color (RGBColor or hex string), or None for transparent fill.
        border_color: Border stroke color (RGBColor or hex string), or None for no border.
        border_width: Border line width in points (default: 0.5 pt).
        corner_radius: Target radius in inches for corners (default: 0.15 in).

    Returns:
        The created Shape object.
    """
    shape = slide.shapes.add_shape(
        MSO_SHAPE.ROUNDED_RECTANGLE,
        _to_inches(left),
        _to_inches(top),
        _to_inches(width),
        _to_inches(height),
    )

    # PowerPoint adjustments[0] is the corner radius fraction relative to min(width, height)
    min_dim = min(width, height)
    if min_dim > 0 and len(shape.adjustments) > 0:
        fraction = corner_radius / min_dim
        # Clamped between 0.0 (sharp rectangle) and 0.5 (full pill shape)
        shape.adjustments[0] = min(0.5, max(0.0, fraction))

    # Fill settings
    if fill_color is not None:
        shape.fill.solid()
        shape.fill.fore_color.rgb = _ensure_rgb(fill_color)
    else:
        shape.fill.background()

    # Border settings
    if border_color is not None:
        shape.line.color.rgb = _ensure_rgb(border_color)
        shape.line.width = Pt(border_width)
    else:
        shape.line.fill.background()

    return shape


def add_text_box(
    slide: Slide,
    left: float,
    top: float,
    width: float,
    height: float,
    text: str,
    font_name: str = "Calibri",
    font_size: int = 12,
    font_color: Union[RGBColor, str, None] = None,
    bold: bool = False,
    italic: bool = False,
    alignment: PP_ALIGN = PP_ALIGN.LEFT,
    word_wrap: bool = True,
    vertical_anchor: MSO_ANCHOR = MSO_ANCHOR.TOP,
    margin_left: float = 0.0,
    margin_top: float = 0.0,
    margin_right: float = 0.0,
    margin_bottom: float = 0.0,
) -> Shape:
    """Creates a text box with the given text and formatting.

    Args:
        slide: The slide to add the text box to.
        left: Left position in inches.
        top: Top position in inches.
        width: Width in inches.
        height: Height in inches.
        text: Text string to display (supports multiline strings).
        font_name: Font family name (default: 'Calibri').
        font_size: Font size in points (default: 12 pt).
        font_color: Text color (RGBColor or hex string).
        bold: Whether text should be bold (default: False).
        italic: Whether text should be italicized (default: False).
        alignment: Text alignment (e.g. PP_ALIGN.LEFT, PP_ALIGN.CENTER).
        word_wrap: Whether text wraps inside the box (default: True).
        vertical_anchor: Vertical alignment (MSO_ANCHOR.TOP, MIDDLE, BOTTOM).
        margin_left: Left margin in inches (default: 0.0).
        margin_top: Top margin in inches (default: 0.0).
        margin_right: Right margin in inches (default: 0.0).
        margin_bottom: Bottom margin in inches (default: 0.0).

    Returns:
        The created Shape object.
    """
    shape = slide.shapes.add_textbox(
        _to_inches(left),
        _to_inches(top),
        _to_inches(width),
        _to_inches(height),
    )
    tf = shape.text_frame
    tf.word_wrap = word_wrap
    tf.vertical_anchor = vertical_anchor
    tf.margin_left = _to_inches(margin_left)
    tf.margin_top = _to_inches(margin_top)
    tf.margin_right = _to_inches(margin_right)
    tf.margin_bottom = _to_inches(margin_bottom)

    lines = text.split("\n") if text else [""]
    resolved_color = _ensure_rgb(font_color)

    for i, line in enumerate(lines):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.text = line
        p.alignment = alignment
        p.font.name = font_name
        if font_size is not None:
            p.font.size = Pt(font_size)
        if resolved_color is not None:
            p.font.color.rgb = resolved_color
        p.font.bold = bold
        p.font.italic = italic

    return shape


def add_styled_text_box(
    slide: Slide,
    left: float,
    top: float,
    width: float,
    height: float,
    runs: list[dict[str, Any]],
    alignment: PP_ALIGN = PP_ALIGN.LEFT,
    word_wrap: bool = True,
    vertical_anchor: MSO_ANCHOR = MSO_ANCHOR.TOP,
    margin_left: float = 0.0,
    margin_top: float = 0.0,
    margin_right: float = 0.0,
    margin_bottom: float = 0.0,
) -> Shape:
    """Creates a text box with multiple styled text runs for mixed formatting.

    Args:
        slide: The slide to add the styled text box to.
        left: Left position in inches.
        top: Top position in inches.
        width: Width in inches.
        height: Height in inches.
        runs: List of run dictionaries. Each run can specify:
              - "text": str (the text to display)
              - "font_name": str (optional font family)
              - "font_size": int (optional font size in points)
              - "font_color": RGBColor or hex str (optional text color)
              - "bold": bool (optional bold flag)
              - "italic": bool (optional italic flag)
              - "new_paragraph": bool (optional flag to begin a new paragraph)
              - "alignment": PP_ALIGN (optional override alignment for new paragraph)
        alignment: Default text alignment for paragraphs (default: PP_ALIGN.LEFT).
        word_wrap: Whether text wraps inside the box (default: True).
        vertical_anchor: Vertical alignment (default: MSO_ANCHOR.TOP).
        margin_left: Left margin in inches (default: 0.0).
        margin_top: Top margin in inches (default: 0.0).
        margin_right: Right margin in inches (default: 0.0).
        margin_bottom: Bottom margin in inches (default: 0.0).

    Returns:
        The created Shape object.
    """
    shape = slide.shapes.add_textbox(
        _to_inches(left),
        _to_inches(top),
        _to_inches(width),
        _to_inches(height),
    )
    tf = shape.text_frame
    tf.word_wrap = word_wrap
    tf.vertical_anchor = vertical_anchor
    tf.margin_left = _to_inches(margin_left)
    tf.margin_top = _to_inches(margin_top)
    tf.margin_right = _to_inches(margin_right)
    tf.margin_bottom = _to_inches(margin_bottom)

    curr_p = tf.paragraphs[0]
    curr_p.alignment = alignment

    for run_data in runs:
        if run_data.get("new_paragraph", False):
            curr_p = tf.add_paragraph()
            curr_p.alignment = run_data.get("alignment", alignment)

        text = str(run_data.get("text", ""))
        lines = text.split("\n")

        for idx, line in enumerate(lines):
            if idx > 0:
                curr_p = tf.add_paragraph()
                curr_p.alignment = run_data.get("alignment", alignment)

            if line:
                r = curr_p.add_run()
                r.text = line
                if run_data.get("font_name"):
                    r.font.name = run_data["font_name"]
                if run_data.get("font_size") is not None:
                    r.font.size = Pt(run_data["font_size"])
                if run_data.get("font_color") is not None:
                    r.font.color.rgb = _ensure_rgb(run_data["font_color"])
                if run_data.get("bold") is not None:
                    r.font.bold = run_data["bold"]
                if run_data.get("italic") is not None:
                    r.font.italic = run_data["italic"]

    return shape


def add_image(
    slide: Slide,
    image_path: str,
    left: float,
    top: float,
    width: Optional[float] = None,
    height: Optional[float] = None,
) -> Picture:
    """Adds an image (SVG, PNG, JPG, WEBP) to the slide.

    For SVG files, it first tries to add directly via python-pptx add_picture.
    If direct insertion fails (due to PIL limitations), it uses a PNG fallback:
    either a pre-existing sibling PNG file or on-the-fly rasterization via CairoSVG.

    All dimension parameters accept float values representing Inches.

    Args:
        slide: The slide to add the image to.
        image_path: Path to the image file.
        left: Left position in inches.
        top: Top position in inches.
        width: Optional width in inches (preserves aspect ratio if height omitted).
        height: Optional height in inches (preserves aspect ratio if width omitted).

    Returns:
        The created Picture shape object.

    Raises:
        FileNotFoundError: If the image file cannot be found.
        RuntimeError: If image loading or SVG conversion fails.
    """
    if not os.path.exists(image_path):
        raise FileNotFoundError(f"Image not found at path: {image_path}")

    l = _to_inches(left)
    t = _to_inches(top)
    w = _to_inches(width)
    h = _to_inches(height)

    kwargs = {}
    if w is not None:
        kwargs["width"] = w
    if h is not None:
        kwargs["height"] = h

    is_svg = image_path.lower().endswith(".svg")

    if is_svg:
        # 1. Attempt direct insertion
        try:
            return slide.shapes.add_picture(image_path, l, t, **kwargs)
        except Exception:
            pass

        # 2. Sibling PNG fallback check (e.g. image.png for image.svg)
        base_no_ext = os.path.splitext(image_path)[0]
        png_sibling = f"{base_no_ext}.png"
        if os.path.isfile(png_sibling):
            try:
                return slide.shapes.add_picture(png_sibling, l, t, **kwargs)
            except Exception:
                pass

        # 3. Dynamic CairoSVG conversion fallback
        try:
            import cairosvg
            png_bytes = cairosvg.svg2png(url=image_path)
            return slide.shapes.add_picture(io.BytesIO(png_bytes), l, t, **kwargs)
        except Exception as err:
            import logging
            logging.getLogger(__name__).warning(
                "CairoSVG conversion not available for '%s' (%s). Using visual placeholder fallback.",
                image_path, err,
            )

        # 4. Graceful Fallback Shape (NEVER crash slide generation if Cairo is missing)
        fb_w = w if w is not None else Inches(0.5)
        fb_h = h if h is not None else Inches(0.5)
        fb_shape = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, l, t, fb_w, fb_h)
        fb_shape.fill.solid()
        fb_shape.fill.fore_color.rgb = RGBColor(0x3B, 0x82, 0xF6)
        fb_shape.line.fill.background()
        if fb_shape.has_text_frame:
            p = fb_shape.text_frame.paragraphs[0]
            p.text = "✦"
            p.font.size = Pt(10)
            p.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
            p.alignment = PP_ALIGN.CENTER
        return fb_shape

    # Standard raster formats (PNG, JPG, etc.)
    return slide.shapes.add_picture(image_path, l, t, **kwargs)


def add_line(
    slide: Slide,
    start_x: float,
    start_y: float,
    end_x: float,
    end_y: float,
    color: Union[RGBColor, str],
    width: float = 1.0,
) -> Connector:
    """Adds a straight line connector between two points.

    Args:
        slide: The slide to add the line to.
        start_x: Starting X position in inches.
        start_y: Starting Y position in inches.
        end_x: Ending X position in inches.
        end_y: Ending Y position in inches.
        color: Line color (RGBColor or hex string).
        width: Line stroke width in points (default: 1.0 pt).

    Returns:
        The created Connector shape object.
    """
    connector = slide.shapes.add_connector(
        MSO_CONNECTOR.STRAIGHT,
        _to_inches(start_x),
        _to_inches(start_y),
        _to_inches(end_x),
        _to_inches(end_y),
    )
    connector.line.color.rgb = _ensure_rgb(color)
    connector.line.width = Pt(width)
    return connector


def add_badge(
    slide: Slide,
    left: float,
    top: float,
    text: str,
    bg_color: Union[RGBColor, str],
    text_color: Union[RGBColor, str],
    font_size: int = 9,
    font_name: str = "Calibri",
    bold: bool = True,
    height: float = 0.3,
) -> Shape:
    """Creates a small pill-shaped badge (narrow rounded rectangle with text).

    Auto-calculates width based on text length. Height is fixed at ~0.3 inches.

    Args:
        slide: The slide to add the badge to.
        left: Left position in inches.
        top: Top position in inches.
        text: Badge label text.
        bg_color: Background pill color (RGBColor or hex string).
        text_color: Text color (RGBColor or hex string).
        font_size: Font size in points (default: 9 pt).
        font_name: Font family name (default: 'Calibri').
        bold: Whether badge text is bold (default: True).
        height: Height of the badge in inches (default: 0.3 in).

    Returns:
        The created Shape object.
    """
    calc_width = max(0.65, len(text) * (font_size * 0.008) + 0.35)

    badge = slide.shapes.add_shape(
        MSO_SHAPE.ROUNDED_RECTANGLE,
        _to_inches(left),
        _to_inches(top),
        _to_inches(calc_width),
        _to_inches(height),
    )
    if len(badge.adjustments) > 0:
        badge.adjustments[0] = 0.5  # Full pill semicircle ends

    badge.fill.solid()
    badge.fill.fore_color.rgb = _ensure_rgb(bg_color)
    badge.line.fill.background()

    tf = badge.text_frame
    tf.word_wrap = False
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    tf.margin_left = Inches(0.08)
    tf.margin_right = Inches(0.08)
    tf.margin_top = Inches(0.02)
    tf.margin_bottom = Inches(0.02)

    p = tf.paragraphs[0]
    p.text = text
    p.alignment = PP_ALIGN.CENTER
    p.font.name = font_name
    p.font.size = Pt(font_size)
    p.font.bold = bold
    p.font.color.rgb = _ensure_rgb(text_color)

    return badge


def add_accent_bar(
    slide: Slide,
    left: float,
    top: float,
    height: float,
    color: Union[RGBColor, str],
    width: float = 0.06,
    rounded: bool = True,
) -> Shape:
    """Adds a thin vertical accent bar (typically used on the left side of cards).

    Args:
        slide: The slide to add the accent bar to.
        left: Left position in inches.
        top: Top position in inches.
        height: Height in inches.
        color: Accent bar color (RGBColor or hex string).
        width: Width in inches (default: 0.06 in).
        rounded: Whether corners are softly rounded (default: True).

    Returns:
        The created Shape object.
    """
    shape_type = MSO_SHAPE.ROUNDED_RECTANGLE if rounded else MSO_SHAPE.RECTANGLE
    bar = slide.shapes.add_shape(
        shape_type,
        _to_inches(left),
        _to_inches(top),
        _to_inches(width),
        _to_inches(height),
    )
    if rounded and len(bar.adjustments) > 0:
        bar.adjustments[0] = 0.5

    bar.fill.solid()
    bar.fill.fore_color.rgb = _ensure_rgb(color)
    bar.line.fill.background()
    return bar
