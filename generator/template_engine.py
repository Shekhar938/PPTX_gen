"""
template_engine.py - PowerPoint Template Infill Engine

Allows loading existing .pptx templates from the templates/ directory or user uploads,
inspecting their text containers and shapes, and populating them with synthesized
content while preserving 100% of the original master layout, colors, and branding.
"""

from __future__ import annotations

import os
from pathlib import Path
from typing import Any, Dict, List, Optional
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor

TEMPLATES_DIR = Path(__file__).resolve().parent.parent / "templates"


def list_available_templates() -> List[Dict[str, Any]]:
    """Scan the templates/ directory and return all available .pptx templates."""
    if not TEMPLATES_DIR.exists():
        TEMPLATES_DIR.mkdir(parents=True, exist_ok=True)

    templates = []
    for file in sorted(TEMPLATES_DIR.glob("*.pptx")):
        try:
            prs = Presentation(str(file))
            templates.append({
                "name": file.stem.replace("_", " ").title(),
                "filename": file.name,
                "path": str(file),
                "slide_count": len(prs.slides),
            })
        except Exception:
            continue
    return templates


def inspect_template_structure(template_path: str) -> List[Dict[str, Any]]:
    """Inspect text frames and placeholders in each slide of a template."""
    prs = Presentation(template_path)
    structure = []

    for idx, slide in enumerate(prs.slides):
        text_frames = []
        for shape in slide.shapes:
            if shape.has_text_frame:
                paragraphs = [p.text.strip() for p in shape.text_frame.paragraphs if p.text.strip()]
                text_frames.append({
                    "shape_id": shape.shape_id,
                    "shape_name": shape.name,
                    "preview": " | ".join(paragraphs[:2]) if paragraphs else "(Empty Box)",
                })
        structure.append({
            "slide_number": idx + 1,
            "placeholders": text_frames,
        })
    return structure


def apply_content_to_template(
    template_path: str,
    slides_content: List[Dict[str, Any]],
    output_path: str,
) -> Dict[str, Any]:
    """Populate an existing PowerPoint template with synthesized content.

    Preserves existing shapes, master layouts, backgrounds, and fonts.
    Replaces placeholder texts in-place.

    Args:
        template_path: Path to the .pptx template to clone.
        slides_content: List of dicts, each with:
            - 'title': str
            - 'subtitle': Optional[str]
            - 'cards': Optional[List[Dict[str, str]]] (with 'title', 'body')
            - 'bullets': Optional[List[str]]
        output_path: Where to save the generated presentation.

    Returns:
        Dict with success status, output_path, and slide_count.
    """
    prs = Presentation(template_path)
    total_template_slides = len(prs.slides)

    for slide_idx, content in enumerate(slides_content):
        # Target the corresponding template slide (or loop/reuse layout if more content than template slides)
        t_idx = slide_idx % total_template_slides
        slide = prs.slides[t_idx]

        # Gather all text-containing shapes on the slide sorted top-to-bottom
        text_shapes = [s for s in slide.shapes if s.has_text_frame]
        text_shapes.sort(key=lambda s: s.top)

        if not text_shapes:
            continue

        # 1. Update Title (usually the top-most text frame)
        title_text = content.get("title", "")
        if title_text and len(text_shapes) > 0:
            title_shape = text_shapes[0]
            if title_shape.text_frame.paragraphs:
                p = title_shape.text_frame.paragraphs[0]
                p.text = title_text

        # 2. Update Subtitle or Description (second top-most shape if on title slide)
        subtitle_text = content.get("subtitle", "")
        if subtitle_text and len(text_shapes) > 1:
            sub_shape = text_shapes[1]
            if sub_shape.text_frame.paragraphs:
                p = sub_shape.text_frame.paragraphs[0]
                p.text = subtitle_text

        # 3. Update Cards / Multi-column content
        cards = content.get("cards", [])
        if cards:
            # Card body shapes are typically shapes 1..N on content slides
            card_shapes = text_shapes[1:] if len(text_shapes) > 1 else text_shapes
            for c_idx, card_data in enumerate(cards):
                if c_idx < len(card_shapes):
                    target_shape = card_shapes[c_idx]
                    tf = target_shape.text_frame
                    c_title = card_data.get("title", "")
                    c_body = card_data.get("body", "")

                    if tf.paragraphs:
                        tf.paragraphs[0].text = c_title
                        if len(tf.paragraphs) > 1:
                            tf.paragraphs[1].text = c_body
                        elif c_body:
                            p = tf.add_paragraph()
                            p.text = c_body
                            p.font.size = Pt(12)

        # 4. Update Bullet Points
        bullets = content.get("bullets", [])
        if bullets and len(text_shapes) > 1:
            body_shape = text_shapes[1]
            tf = body_shape.text_frame
            tf.clear()
            for b_idx, bullet in enumerate(bullets):
                p = tf.paragraphs[0] if b_idx == 0 else tf.add_paragraph()
                p.text = bullet

    os.makedirs(os.path.dirname(os.path.abspath(output_path)), exist_ok=True)
    prs.save(output_path)

    return {
        "success": True,
        "output_path": os.path.abspath(output_path),
        "slide_count": len(prs.slides),
        "error": None,
    }
