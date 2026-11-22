"""
gemini_director.py - LLM-Powered Creative Director for Slide Generation

This module uses Google Gemini to:
1. Analyze and synthesize content from multiple sources (YouTube transcripts,
   PDFs, DOCX files, existing PPTX files).
2. Generate Python code that uses the SlideBuilder API to create each slide.
3. Execute the generated code to produce the final presentation.
"""

from __future__ import annotations

import logging
import os
import re
import textwrap
import time
import traceback
from pathlib import Path
from typing import Optional

from google import genai
from google.genai import errors as genai_errors

logger = logging.getLogger(__name__)

# Ordered fallback chain of models available on this API key
MODEL_FALLBACK_CHAIN = [
    "gemini-3.5-flash",
    "gemini-3.5-flash-lite",
    "gemini-3.8-flash",
]

# ---------------------------------------------------------------------------
# System prompt for the LLM "Creative Director"
# ---------------------------------------------------------------------------
SYSTEM_PROMPT = textwrap.dedent("""\
You are an expert presentation designer and creative director. Your job is to
take raw content (YouTube transcripts, document text, user instructions) and
produce a stunning, professional PowerPoint presentation using a Python API.

## YOUR OUTPUT FORMAT

You MUST output ONLY valid Python code that uses the SlideBuilder API.
Do NOT include any explanation, markdown fences, or commentary outside the code.
The code will be executed directly.

## AVAILABLE API

You have access to these pre-imported objects:

```python
deck   # PresentationDeck instance (already created with the user's chosen theme)
```

### Creating a new slide:
```python
slide = deck.new_slide()
```

### Slide components (call these on the slide object):

1. **slide.add_hero_heading(title, subtitle="", badge="", author="")**
   - Full-width title slide. Use for the first slide or section openers.

2. **slide.add_section_header(title, category="")**
   - Section divider with category badge and large title.

3. **slide.add_character(pose, side="left", speech_bubble="", size=2.5)**
   - Places a doodle character.
   - Poses (38+ available): thinking, pointing, celebrating, confused, presenting,
     reading, chilling, dancing, coffee, meditating, jumping, running, sprinting,
     strolling, sitting, sittingreading, unboxing, selfie, plant, levitate, groovy,
     loving, float, doggie, dogjump, clumsy, ballet, icecream, laying, rollerskating, etc.
   - Use speech_bubble for the character to "say" something.

4. **slide.add_bento_cards(cards, columns=0, card_height=0)**
   - Grid of styled cards. Each card is a dict:
     {"title": str, "body": str, "icon": str, "highlight": str}
   - Over 2,400+ Lucide vector icons are available by name, e.g.:
     brain, database, shield, lightning, chart_bar, target, star, arrow_up, clock,
     check_circle, warning, lightbulb, users, code, globe, rocket, server, cloud,
     cpu, zap, lock, key, terminal, git-branch, heart, dollar-sign, percent,
     sparkles, file-text, layers, refresh-cw, compass, search, bookmark, eye, etc.

5. **slide.add_stat_banner(metrics)**
   - Row of big-number callouts. Each metric: {"value": str, "label": str}

6. **slide.add_bullet_list(items, title="", position="full", icon="")**
   - Styled bullet list. position: "full", "left", "right"

7. **slide.add_comparison(left_data, right_data)**
   - Side-by-side comparison. Each side: {"title": str, "points": [str]}

8. **slide.add_timeline(steps)**
   - Horizontal process flow. Each step: {"title": str, "description": str}

9. **slide.add_quote(text, author="")**
   - Large quote callout with optional attribution.

10. **slide.add_chart(chart_type, categories, series, title="", height=3.5)**
    - Native PowerPoint chart. chart_type: "bar", "column", "line", "pie", "area"
    - series: [{"name": str, "values": [numbers]}]

11. **slide.add_table(headers, rows, title="")**
    - Styled table. rows: [[str, str, ...], ...]

12. **slide.add_content_with_image(title, body, image_path="", image_side="right")**
    - Split layout with text and image. image_path can be an icon name.

13. **slide.add_footer(text="", show_page_number=True, page_number=0)**
    - Subtle footer bar.

## DESIGN PRINCIPLES - FOLLOW STRICTLY

1. **NEVER create a boring slide.** Every slide must have visual structure
   (cards, stats, timelines, characters, charts, or comparisons).
2. **Use characters** on at least 2-3 slides to maintain visual personality
   and narrative flow.
3. **Vary the layouts** - don't use the same component on consecutive slides.
   Alternate between bento cards, stat banners, timelines, comparisons, etc.
4. **Extract real data** from the source content. Use actual numbers, quotes,
   and facts. Never make up statistics.
5. **Keep text concise** - maximum 3 lines per card body, 5-7 words per card title.
6. **Use badges/categories** to organize slides into logical sections.
7. **Add footers** with page numbers on every slide (except the title slide).
8. **Create 6-15 slides** depending on content volume.
9. **First slide** should always be a hero_heading with the presentation title.
10. **Last slide** should be a summary or key takeaways slide.

## IMPORTANT RULES

- Output ONLY Python code. No markdown, no explanations.
- Do NOT import anything - all objects are pre-imported.
- Do NOT call deck.save() - that's handled externally.
- Use the ACTUAL content from the source material, not placeholders.
- Ensure slide count matches the depth of the content.
""")


def _build_content_prompt(
    user_prompt: str,
    youtube_data: list[dict] | None = None,
    document_data: list[dict] | None = None,
) -> str:
    """Build the user message combining all source content and instructions."""

    parts: list[str] = []

    parts.append("## USER INSTRUCTIONS")
    parts.append(user_prompt if user_prompt else "Create a comprehensive presentation from the provided content.")
    parts.append("")

    if youtube_data:
        parts.append("## YOUTUBE VIDEO TRANSCRIPTS")
        for i, yt in enumerate(youtube_data, 1):
            parts.append(f"\n### Video {i}: {yt.get('title', yt.get('video_id', 'Unknown'))}")
            parts.append(f"Video ID: {yt.get('video_id', 'N/A')}")
            transcript = yt.get("transcript", "")
            # Truncate very long transcripts to avoid token limits
            if len(transcript) > 15000:
                transcript = transcript[:15000] + "\n\n[... transcript truncated for brevity ...]"
            parts.append(f"Transcript:\n{transcript}")
        parts.append("")

    if document_data:
        parts.append("## DOCUMENT CONTENT")
        for i, doc in enumerate(document_data, 1):
            source = doc.get("source_type", "document").upper()
            parts.append(f"\n### Document {i} ({source})")

            if "sections" in doc and doc["sections"]:
                for sec in doc["sections"][:30]:  # Limit sections
                    heading = sec.get("heading", "")
                    content = sec.get("content", "")
                    if heading:
                        parts.append(f"\n**{heading}**")
                    if content:
                        if len(content) > 3000:
                            content = content[:3000] + "\n[... truncated ...]"
                        parts.append(content)
            elif "slides" in doc and doc["slides"]:
                for sl in doc["slides"][:30]:
                    parts.append(f"\nSlide {sl.get('slide_number', '?')}: {sl.get('title', '')}")
                    parts.append(sl.get("content", ""))
            else:
                text = doc.get("text", "")
                if len(text) > 15000:
                    text = text[:15000] + "\n\n[... truncated ...]"
                parts.append(text)
        parts.append("")

    parts.append("\n## TASK")
    parts.append(
        "Now generate Python code using the SlideBuilder API to create a "
        "stunning presentation from ALL the content above. Follow the design "
        "principles strictly. Output ONLY valid Python code."
    )

    return "\n".join(parts)


class GeminiDirector:
    """Orchestrates the LLM to generate slide-building Python code."""

    def __init__(self, api_key: str, model: str = "gemini-3.5-flash"):
        self.client = genai.Client(api_key=api_key)
        self.model = model

    def generate_slide_code(
        self,
        user_prompt: str,
        youtube_data: list[dict] | None = None,
        document_data: list[dict] | None = None,
        theme: str = "bw_doodle",
        status_callback=None,
    ) -> str:
        """Generate Python code for slide creation.

        Retries with exponential backoff on 503 overload errors, then falls
        back through MODEL_FALLBACK_CHAIN before giving up.
        Reports real-time status to status_callback if provided.
        """
        content_prompt = _build_content_prompt(user_prompt, youtube_data, document_data)

        # Build the full model list to try: requested model first, then fallbacks
        model_queue = [self.model] + [m for m in MODEL_FALLBACK_CHAIN if m != self.model]

        last_exc: Exception | None = None

        for model_name in model_queue:
            for attempt in range(3):
                wait = 2 ** attempt  # 1s, 2s, 4s
                try:
                    msg = f"Connecting to {model_name} (attempt {attempt + 1})..."
                    logger.info(msg)
                    if status_callback:
                        status_callback(msg)

                    response = self.client.models.generate_content(
                        model=model_name,
                        contents=content_prompt,
                        config=genai.types.GenerateContentConfig(
                            system_instruction=SYSTEM_PROMPT,
                            temperature=0.7,
                            max_output_tokens=16000,
                        ),
                    )
                    code = response.text.strip()
                    # Strip markdown code fences if present
                    code = re.sub(r"^```(?:python)?\s*\n", "", code)
                    code = re.sub(r"\n```\s*$", "", code)

                    if model_name != self.model:
                        note = f"Synthesized with {model_name} (fallback from {self.model})"
                        logger.warning(note)
                        if status_callback:
                            status_callback(f"✅ {note}")
                    else:
                        if status_callback:
                            status_callback(f"✅ AI response received from {model_name}!")
                    return code

                except (genai_errors.ServerError, Exception) as exc:
                    is_overload = (
                        isinstance(exc, genai_errors.ServerError)
                        and getattr(exc, "status_code", 0) == 503
                    ) or "503" in str(exc) or "UNAVAILABLE" in str(exc) or "high demand" in str(exc)

                    is_not_found = (
                        "404" in str(exc) or "NOT_FOUND" in str(exc)
                        or "no longer available" in str(exc)
                    )

                    if is_not_found:
                        warn_msg = f"Model {model_name} unavailable on your API key (404). Trying next model..."
                        logger.warning(warn_msg)
                        if status_callback:
                            status_callback(f"⚠️ {warn_msg}")
                        last_exc = exc
                        break

                    if is_overload and attempt < 2:
                        retry_msg = f"{model_name} is in high demand on Google servers. Retrying in {wait}s..."
                        logger.warning(retry_msg)
                        if status_callback:
                            status_callback(f"⏳ {retry_msg}")
                        time.sleep(wait)
                        continue

                    last_exc = exc
                    fail_msg = f"{model_name} failed. Trying fallback model..."
                    logger.warning(fail_msg)
                    if status_callback:
                        status_callback(f"🔄 {fail_msg}")
                    break

        raise last_exc or RuntimeError("All Gemini models in the fallback chain failed.")

    def execute_slide_code(
        self,
        code: str,
        theme: str = "bw_doodle",
        output_path: str = "output.pptx",
    ) -> dict:
        """Execute the generated Python code to build the presentation."""
        from .slide_builder import PresentationDeck

        deck = PresentationDeck(theme=theme)

        # Build execution namespace with auto-string resolution
        class RobustNamespace(dict):
            def __missing__(self, key):
                return key

        namespace = RobustNamespace({
            "deck": deck,
            "PresentationDeck": PresentationDeck,
            "__builtins__": __builtins__,
        })
        for token in [
            "brain", "database", "shield", "lightning", "chart_bar", "target", "star",
            "arrow_up", "clock", "check_circle", "warning", "lightbulb", "users", "code", "globe",
            "thinking", "pointing", "celebrating", "confused", "presenting", "reading",
            "left", "right", "center", "full", "bar", "column", "pie", "line", "area",
        ]:
            namespace[token] = token

        try:
            exec(code, namespace)
            saved_path = deck.save(output_path)
            return {
                "success": True,
                "output_path": saved_path,
                "slide_count": deck.slide_count,
                "error": None,
                "code": code,
            }
        except Exception as e:
            return {
                "success": False,
                "output_path": None,
                "slide_count": 0,
                "error": f"{type(e).__name__}: {e}\n{traceback.format_exc()}",
                "code": code,
            }

    def generate_and_build(
        self,
        user_prompt: str,
        youtube_data: list[dict] | None = None,
        document_data: list[dict] | None = None,
        theme: str = "bw_doodle",
        output_path: str = "output.pptx",
        max_retries: int = 1,
        status_callback=None,
    ) -> dict:
        """Full pipeline: generate code, execute, retry on failure.
        Gracefully catches all exceptions and reports them through status_callback.
        """
        last_error = None

        for attempt in range(max_retries + 1):
            try:
                if attempt == 0:
                    code = self.generate_slide_code(
                        user_prompt, youtube_data, document_data, theme,
                        status_callback=status_callback,
                    )
                else:
                    if status_callback:
                        status_callback("Fixing slide code and re-generating...")
                    retry_prompt = (
                        f"{user_prompt}\n\n"
                        f"## PREVIOUS ATTEMPT FAILED\n"
                        f"Error: {last_error}\n\n"
                        f"Please fix the code and try again. Output ONLY valid Python code."
                    )
                    code = self.generate_slide_code(
                        retry_prompt, youtube_data, document_data, theme,
                        status_callback=status_callback,
                    )

                if status_callback:
                    status_callback("🎨 Assembling PowerPoint presentation with vector shapes...")

                # Execute
                result = self.execute_slide_code(code, theme, output_path)

                if result["success"]:
                    if status_callback:
                        status_callback(f"🎉 Complete! Generated {result['slide_count']} slides.")
                    return result

                last_error = result["error"]
                if status_callback:
                    status_callback(f"⚠️ Code execution error: {last_error.splitlines()[0] if last_error else ''}")

            except Exception as e:
                err_detail = f"{type(e).__name__}: {e}"
                logger.error("Generation error: %s", err_detail)
                if status_callback:
                    status_callback(f"❌ Error: {err_detail}")
                return {
                    "success": False,
                    "output_path": None,
                    "slide_count": 0,
                    "error": err_detail,
                    "code": "",
                }

        return result  # Return last failed result
